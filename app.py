import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import networkx as nx
from sqlalchemy import text
import io
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ============================================================
# SAYFA AYARLARI VE CSS (Kırmızı/Lacivert Kurumsal Tema)
# ============================================================
st.set_page_config(page_title="Emir Asel İnşaat — Proje Yönetimi", page_icon="🏗️", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #F8F9FA; }
    .header-container {
        background: linear-gradient(135deg, #0A192F 0%, #112240 100%);
        padding: 20px 30px; border-radius: 15px; color: white; margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .header-title { color: white !important; font-size: 28px !important; font-weight: 800 !important; margin: 0; letter-spacing: 1px; }
    div[data-testid="stMetric"] {
        background-color: white; border-radius: 10px; padding: 15px 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05); border-left: 5px solid #E63946;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    div[data-testid="stMetric"]:hover { transform: translateY(-3px); box-shadow: 0 5px 15px rgba(230, 57, 70, 0.2); }
    div[data-testid="stDataFrame"] { border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); background-color: white; padding: 10px; }
    section[data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #EAEAEA; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# VERİTABANI BAĞLANTISI VE MIGRATION
# ============================================================
@st.cache_resource
def get_db_connection():
    try:
        conn = st.connection("neon", type="sql")
        return conn
    except Exception as e:
        st.error(f"Veritabanı bağlantı hatası: {e}")
        st.info("Lütfen '.streamlit/secrets.toml' dosyanızda 'neon' bağlantı ayarlarını yapılandırın.")
        st.stop()

conn = get_db_connection()

def init_db():
    try:
        with conn.session as s:
            # Mevcut Tablolar
            s.execute(text("CREATE TABLE IF NOT EXISTS ilerleme (wbs_kodu TEXT PRIMARY KEY, ilerleme_yuzdesi INTEGER DEFAULT 0)"))
            s.execute(text("CREATE TABLE IF NOT EXISTS satin_alma (id SERIAL PRIMARY KEY, wbs_kodu TEXT, malzeme TEXT, tedarikci TEXT, tutar NUMERIC, durum TEXT, tarih TEXT)"))
            s.execute(text("CREATE TABLE IF NOT EXISTS iscilik (id SERIAL PRIMARY KEY, wbs_kodu TEXT, ekip TEXT, is_tanimi TEXT, adet INTEGER, yevmiye NUMERIC, tarih TEXT)"))
            
            # Yeni / Güncellenen Tablolar
            s.execute(text("""
                CREATE TABLE IF NOT EXISTS activities (
                    wbs_code TEXT PRIMARY KEY, task_name TEXT, duration INTEGER, budget NUMERIC,
                    responsible TEXT, plan_qty NUMERIC DEFAULT 0, act_qty NUMERIC DEFAULT 0, unit TEXT,
                    manual_progress NUMERIC DEFAULT 0
                )
            """))
            s.execute(text("CREATE TABLE IF NOT EXISTS dependencies (id SERIAL PRIMARY KEY, predecessor TEXT, successor TEXT, relation_type TEXT DEFAULT 'FS')"))
            s.execute(text("""
                CREATE TABLE IF NOT EXISTS progress_history (
                    id SERIAL PRIMARY KEY, wbs_code TEXT, log_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    progress NUMERIC, note TEXT
                )
            """))
            s.commit()
    except Exception as e:
        st.warning(f"Veritabanı tablo kontrolü uyarısı: {e}")

init_db()

# ============================================================
# MÜHENDİSLİK FONKSİYONLARI (CPM & EVM)
# ============================================================
def calculate_cpm(act_df, dep_df):
    if act_df.empty: return pd.DataFrame()
    G = nx.DiGraph()
    for _, row in act_df.iterrows(): G.add_node(row['wbs_code'])
    for _, row in dep_df.iterrows(): G.add_edge(row['predecessor'], row['successor'])
    if not nx.is_directed_acyclic_graph(G):
        st.error("⚠️ HATA: Döngüsel bağımlılık (A->B->C->A) tespit edildi!")
        return pd.DataFrame()

    topo = list(nx.topological_sort(G))
    es, ef, ls, lf = {}, {}, {}, {}

    # İleri Geçiş
    for node in topo:
        act = act_df[act_df['wbs_code'] == node].iloc[0]
        dur = act['duration']
        preds = dep_df[dep_df['successor'] == node]
        if preds.empty: es[node] = 1
        else:
            max_s = 1
            for _, p in preds.iterrows():
                p_wbs, rel = p['predecessor'], p['relation_type']
                if rel == 'FS': max_s = max(max_s, ef.get(p_wbs, 1) + 1)
                elif rel == 'SS': max_s = max(max_s, es.get(p_wbs, 1))
                elif rel == 'FF': max_s = max(max_s, ef.get(p_wbs, 1) - dur + 1)
            es[node] = max(1, max_s)
        ef[node] = es[node] + dur - 1

    proj_bitis = max(ef.values()) if ef else 1

    # Geri Geçiş
    for node in reversed(topo):
        act = act_df[act_df['wbs_code'] == node].iloc[0]
        dur = act['duration']
        succs = dep_df[dep_df['predecessor'] == node]
        if succs.empty: lf[node] = proj_bitis
        else:
            min_f = float('inf')
            for _, s in succs.iterrows():
                s_wbs, rel = s['successor'], s['relation_type']
                if rel == 'FS': min_f = min(min_f, ls.get(s_wbs, proj_bitis) - 1)
                elif rel == 'SS': min_f = min(min_f, ls.get(s_wbs, proj_bitis) + dur - 1)
                elif rel == 'FF': min_f = min(min_f, lf.get(s_wbs, proj_bitis))
            lf[node] = min_f
        ls[node] = lf[node] - dur + 1

    res = []
    for _, row in act_df.iterrows():
        wbs = row['wbs_code']
        slack = ls[wbs] - es[wbs]
        # Metraj İlerlemesi (%)
        qty_prog = (row['act_qty'] / row['plan_qty'] * 100) if row['plan_qty'] > 0 else 0
        # Karma İlerleme (%30 Manuel, %70 Metraj)
        final_prog = (row['manual_progress'] * 0.3) + (qty_prog * 0.7)
        
        res.append({
            "WBS": wbs, "İş Kalemi": row['task_name'], "Süre": row['duration'],
            "Bütçe": row['budget'], "ES": es[wbs], "EF": ef[wbs],
            "Bolluk": slack, "Kritik": "EVET" if slack == 0 else "HAYIR",
            "Fiziksel İlerleme (%)": round(final_prog, 1)
        })
    return pd.DataFrame(res)

def calculate_evm(cpm_df, ac_total):
    if cpm_df.empty: return {}
    
    total_budget = cpm_df['Bütçe'].sum()
    # Ağırlıklı İlerleme
    ev = sum((row['Bütçe'] * row['Fiziksel İlerleme (%)'] / 100) for _, row in cpm_df.iterrows())
    
    # Planlanan Değer (PV) - Basit Time-Based S-Curve (Şimdilik proje süresinin %50'sindeyiz varsayımı)
    # Gerçekte takvim tarihine göre hesaplanır.
    pv = total_budget * 0.5 
    
    spi = ev / pv if pv > 0 else 0
    cpi = ev / ac_total if ac_total > 0 else 0
    eac = total_budget / cpi if cpi > 0 else total_budget
    
    return {"PV": pv, "EV": ev, "AC": ac_total, "SPI": round(spi,2), "CPI": round(cpi,2), "EAC": eac, "BAC": total_budget}

# ============================================================
# ANA UYGULAMA
# ============================================================
st.markdown('<div class="header-container"><h1 class="header-title">🏗️ EMİR ASEL İNŞAAT — PROJE YÖNETİM PANELİ</h1></div>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Menü")
    menu = st.radio("Sayfa Seçin", [
        "📊 Dashboard & Raporlar", 
        "📅 İş Programı & CPM", 
        "📈 İlerleme Takibi", 
        "🛒 Satın Alma", 
        "👷 İşçilik & Puantaj"
    ])

# Verileri Çek
df_act = conn.query("SELECT * FROM activities", ttl=0)
df_dep = conn.query("SELECT * FROM dependencies", ttl=0)
df_satinalma = conn.query("SELECT * FROM satin_alma", ttl=0)
df_iscilik = conn.query("SELECT * FROM iscilik", ttl=0)
cpm_df = calculate_cpm(df_act, df_dep) if not df_act.empty else pd.DataFrame()

# ============================================================
# SAYFA 1: DASHBOARD & RAPORLAR
# ============================================================
if menu == "📊 Dashboard & Raporlar":
    st.subheader("Yönetici Özeti")
    
    if df_act.empty:
        st.info("Lütfen önce 'İş Programı' sekmesinden iş kalemlerini ekleyin.")
    else:
        # Gerçekleşen Maliyet (AC)
        ac_satinalma = df_satinalma['tutar'].sum() if not df_satinalma.empty else 0
        ac_iscilik = (df_iscilik['adet'] * df_iscilik['yevmiye']).sum() if not df_iscilik.empty else 0
        ac_total = ac_satinalma + ac_iscilik

        evm = calculate_evm(cpm_df, ac_total)
        avg_prog = cpm_df['Fiziksel İlerleme (%)'].mean()

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Toplam Bütçe (BAC)", f"{evm['BAC']:,.0f} ₺")
        col2.metric("Gerçekleşen Maliyet (AC)", f"{evm['AC']:,.0f} ₺")
        col3.metric("Kazanılmış Değer (EV)", f"{evm['EV']:,.0f} ₺")
        col4.metric("Fiziksel İlerleme", f"%{avg_prog:.1f}")

        col5, col6, col7, col8 = st.columns(4)
        col5.metric("Maliyet Performansı (CPI)", evm['CPI'], help="1.0'dan küçükse bütçe aşımı var.")
        col6.metric("Zaman Performansı (SPI)", evm['SPI'], help="1.0'dan küçükse proje geride.")
        col7.metric("Tahmini Bitiş Maliyeti (EAC)", f"{evm['EAC']:,.0f} ₺")
        col8.metric("Kalan Maliyet (ETC)", f"{evm['EAC'] - evm['AC']:,.0f} ₺")

        st.markdown("---")
        
        # S-Curve ve Grafikler
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Maliyet Dağılımı")
            fig_pie = px.pie(values=[ac_satinalma, ac_iscilik], names=["Satın Alma", "İşçilik"], 
                         color_discrete_sequence=['#E63946', '#0A192F'], hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with c2:
            st.markdown("#### S-Eğrisi (Planlanan vs Gerçekleşen)")
            # Simüle edilmiş S-Curve
            days = np.arange(0, 100, 10)
            plan_curve = 1 / (1 + np.exp(-0.1 * (days - 50))) * 100
            actual_curve = plan_curve * (avg_prog / 50) # Basit simülasyon
            
            fig_scurve = go.Figure()
            fig_scurve.add_trace(go.Scatter(x=days, y=plan_curve, mode='lines', name='Planlanan', line=dict(color='#0A192F', width=3)))
            fig_scurve.add_trace(go.Scatter(x=days, y=actual_curve, mode='lines', name='Gerçekleşen', line=dict(color='#E63946', width=3, dash='dash')))
            fig_scurve.update_layout(xaxis_title="Proje Süresi (%)", yaxis_title="İlerleme (%)")
            st.plotly_chart(fig_scurve, use_container_width=True)

        # PDF Raporlama
        st.markdown("---")
        st.subheader("📄 Yönetici Raporu")
        if st.button("PDF Raporu Oluştur ve İndir"):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            styles = getSampleStyleSheet()
            elements = []
            
            # PDF Başlık
            title_style = ParagraphStyle('Title', parent=styles['Heading1'], textColor=colors.HexColor('#E63946'))
            elements.append(Paragraph("EMİR ASEL İNŞAAT - PROJE YÖNETİM RAPORU", title_style))
            elements.append(Spacer(1, 12))
            
            # PDF Tablo
            data = [
                ['Metrik', 'Değer'],
                ['Toplam Bütçe (BAC)', f"{evm['BAC']:,.2f} TL"],
                ['Gerçekleşen Maliyet (AC)', f"{evm['AC']:,.2f} TL"],
                ['Kazanılmış Değer (EV)', f"{evm['EV']:,.2f} TL"],
                ['Maliyet Performansı (CPI)', evm['CPI']],
                ['Zaman Performansı (SPI)', evm['SPI']],
                ['Fiziksel İlerleme', f"%{avg_prog:.1f}"],
                ['Tahmini Bitiş Maliyeti (EAC)', f"{evm['EAC']:,.2f} TL"],
            ]
            table = Table(data, colWidths=[200, 200])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0A192F')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            elements.append(table)
            doc.build(elements)
            
            st.download_button(
                label="📥 PDF Raporunu İndir",
                data=buffer.getvalue(),
                file_name="EmirAsel_Yonetici_Raporu.pdf",
                mime="application/pdf"
            )

# ============================================================
# SAYFA 2: İŞ PROGRAMI & CPM
# ============================================================
elif menu == "📅 İş Programı & CPM":
    st.subheader("WBS ve Kritik Yol (CPM)")
    tab1, tab2, tab3 = st.tabs(["Aktiviteler", "Bağımlılıklar", "CPM & Gantt"])
    
    with tab1:
        with st.form("add_activity"):
            c1, c2, c3, c4 = st.columns(4)
            wbs = c1.text_input("WBS Kodu (Örn: A01)")
            name = c2.text_input("İş Kalemi Adı")
            dur = c3.number_input("Süre (Gün)", min_value=1, value=1)
            bud = c4.number_input("Bütçe (TL)", min_value=0.0, value=0.0)
            if st.form_submit_button("Ekle / Güncelle"):
                with conn.session as s:
                    s.execute(text("""
                        INSERT INTO activities (wbs_code, task_name, duration, budget)
                        VALUES (:w, :n, :d, :b)
                        ON CONFLICT (wbs_code) DO UPDATE SET task_name=:n, duration=:d, budget=:b
                    """), {"w": wbs, "n": name, "d": dur, "b": bud})
                    s.commit()
                st.rerun()
        st.dataframe(df_act, use_container_width=True)

    with tab2:
        with st.form("add_dep"):
            c1, c2, c3 = st.columns(3)
            pred = c1.text_input("Öncül WBS")
            succ = c2.text_input("Ardıl WBS")
            rel = c3.selectbox("İlişki Tipi", ["FS", "SS", "FF", "SF"])
            if st.form_submit_button("Bağımlılık Ekle"):
                with conn.session as s:
                    s.execute(text("INSERT INTO dependencies (predecessor, successor, relation_type) VALUES (:p, :s, :r)"), 
                              {"p": pred, "s": succ, "r": rel})
                    s.commit()
                st.rerun()
        st.dataframe(df_dep, use_container_width=True)

    with tab3:
        if not cpm_df.empty:
            st.dataframe(cpm_df.style.applymap(lambda x: 'background-color: #ffcccc' if x == "EVET" else '', subset=['Kritik']), use_container_width=True)
            
            # Gantt
            start_date = datetime(2026, 1, 15)
            gantt_data = []
            for _, row in cpm_df.iterrows():
                gantt_data.append(dict(
                    Task=row['İş Kalemi'],
                    Start=start_date + timedelta(days=int(row['ES'])-1),
                    Finish=start_date + timedelta(days=int(row['EF'])-1),
                    Durum="Kritik" if row['Kritik'] == "EVET" else "Normal"
                ))
            fig = px.timeline(pd.DataFrame(gantt_data), x_start="Start", x_end="Finish", y="Task", color="Durum",
                              color_discrete_map={"Kritik": "#E63946", "Normal": "#0A192F"})
            fig.update_yaxes(autorange="reversed")
            st.plotly_chart(fig, use_container_width=True)

# ============================================================
# SAYFA 3: İLERLEME TAKİBİ (Metraj + Manuel)
# ============================================================
elif menu == "📈 İlerleme Takibi":
    st.subheader("İlerleme ve Metraj Girişi")
    
    if df_act.empty:
        st.warning("Önce iş kalemi ekleyin.")
    else:
        with st.form("progress_form"):
            st.markdown("#### İmalat ve Metraj Güncellemesi")
            selected_wbs = st.selectbox("İş Kalemi", df_act['wbs_code'] + " - " + df_act['task_name'])
            wbs_code = selected_wbs.split(" - ")[0]
            
            act_data = df_act[df_act['wbs_code'] == wbs_code].iloc[0]
            
            c1, c2 = st.columns(2)
            plan_qty = c1.number_input("Planlanan Miktar (Örn: Ton, m3)", value=float(act_data['plan_qty']))
            act_qty = c2.number_input("Gerçekleşen Miktar", value=float(act_data['act_qty']))
            
            man_prog = st.slider("Manuel İlerleme (%)", 0, 100, int(act_data['manual_progress']))
            
            if st.form_submit_button("Kaydet ve Geçmişe İşle"):
                with conn.session as s:
                    # Aktiviteleri güncelle
                    s.execute(text("UPDATE activities SET plan_qty=:pq, act_qty=:aq, manual_progress=:mp WHERE wbs_code=:w"), 
                              {"pq": plan_qty, "aq": act_qty, "mp": man_prog, "w": wbs_code})
                    # Geçmişe logla
                    s.execute(text("INSERT INTO progress_history (wbs_code, progress, note) VALUES (:w, :p, :n)"),
                              {"w": wbs_code, "p": man_prog, "n": f"Metraj: {act_qty}/{plan_qty}"})
                    s.commit()
                st.success("İlerleme tarihçeye kaydedildi.")
                st.rerun()
                
        st.markdown("---")
        st.subheader("Son İlerleme Hareketleri")
        df_hist = conn.query("SELECT * FROM progress_history ORDER BY log_date DESC LIMIT 10", ttl=0)
        st.dataframe(df_hist, use_container_width=True)

# ============================================================
# SAYFA 4: SATIN ALMA
# ============================================================
elif menu == "🛒 Satın Alma":
    st.subheader("Satın Alma (Gerçek Maliyet Girişi)")
    
    with st.form("satin_alma_form"):
        c1, c2, c3, c4 = st.columns(4)
        wbs_sat = c1.selectbox("İlgili WBS Kodu (Maliyet Dağıtımı İçin)", df_act['wbs_code'] if not df_act.empty else ["Yok"])
        malzeme = c2.text_input("Malzeme")
        tutar = c3.number_input("Tutar (TL)", min_value=0.0)
        durum = c4.selectbox("Durum", ["Sipariş Verildi", "Teslim Edildi", "Ödendi"])
        
        if st.form_submit_button("Kaydet"):
            with conn.session as s:
                s.execute(text("""
                    INSERT INTO satin_alma (wbs_kodu, malzeme, tutar, durum, tarih) 
                    VALUES (:w, :m, :t, :d, :tr)
                """), {"w": wbs_sat, "m": malzeme, "t": tutar, "d": durum, "tr": datetime.now().strftime("%Y-%m-%d")})
                s.commit()
            st.rerun()
    st.dataframe(df_satinalma, use_container_width=True)

# ============================================================
# SAYFA 5: İŞÇİLİK & PUANTAJ
# ============================================================
elif menu == "👷 İşçilik & Puantaj":
    st.subheader("İşçilik & Puantaj")
    
    with st.form("iscilik_form"):
        c1, c2, c3, c4 = st.columns(4)
        wbs_isc = c1.selectbox("İlgili WBS Kodu", df_act['wbs_code'] if not df_act.empty else ["Yok"])
        ekip = c2.text_input("Ekip Adı")
        adet = c3.number_input("Kişi Sayısı", min_value=1)
        yevmiye = c4.number_input("Günlük Yevmiye (TL)", min_value=0.0)
        
        if st.form_submit_button("Kaydet"):
            with conn.session as s:
                s.execute(text("""
                    INSERT INTO iscilik (wbs_kodu, ekip, adet, yevmiye, tarih) 
                    VALUES (:w, :e, :a, :y, :t)
                """), {"w": wbs_isc, "e": ekip, "a": adet, "y": yevmiye, "t": datetime.now().strftime("%Y-%m-%d")})
                s.commit()
            st.rerun()
    st.dataframe(df_iscilik, use_container_width=True)
