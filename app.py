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
# 1. SAYFA AYARLARI VE KURUMSAL TEMA (CSS)
# ============================================================
st.set_page_config(page_title="Emir Asel İnşaat — Proje Yönetimi", page_icon="🏗️", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #F4F7F6; }
    .header-container {
        background: linear-gradient(135deg, #0A192F 0%, #112240 100%);
        padding: 25px 30px; border-radius: 15px; color: white; margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .header-title { color: white !important; font-size: 28px !important; font-weight: 800 !important; margin: 0; letter-spacing: 1px; }
    div[data-testid="stMetric"] {
        background-color: white; border-radius: 12px; padding: 15px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04); border-left: 6px solid #E63946;
        transition: transform 0.2s ease;
    }
    div[data-testid="stMetric"]:hover { transform: translateY(-3px); box-shadow: 0 6px 15px rgba(230, 57, 70, 0.15); }
    div[data-testid="stDataFrame"] { border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); background-color: white; padding: 10px; }
    section[data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #EAEAEA; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 2. VERİTABANI BAĞLANTISI VE MIGRATION (Veri Kaybı Olmadan)
# ============================================================
@st.cache_resource
def get_db_connection():
    try:
        return st.connection("neon", type="sql")
    except Exception as e:
        st.error(f"Veritabanı bağlantı hatası: {e}")
        st.info("Lütfen '.streamlit/secrets.toml' dosyanızda 'neon' bağlantı ayarlarını yapılandırın.")
        st.stop()

conn = get_db_connection()

def init_db():
    """Mevcut tabloları SİLMEDEN yeni profesyonel tabloları ve sütunları ekler."""
    try:
        with conn.session as s:
            # 1. ESKİ TABLOLAR (Korunuyor) ve WBS Kodlarının Eklenmesi
            s.execute(text("CREATE TABLE IF NOT EXISTS ilerleme (wbs_kodu TEXT PRIMARY KEY, ilerleme_yuzdesi INTEGER DEFAULT 0)"))
            s.execute(text("CREATE TABLE IF NOT EXISTS satin_alma (id SERIAL PRIMARY KEY, malzeme TEXT, tedarikci TEXT, tutar NUMERIC, durum TEXT, tarih TEXT)"))
            s.execute(text("CREATE TABLE IF NOT EXISTS iscilik (id SERIAL PRIMARY KEY, ekip TEXT, is_tanimi TEXT, adet INTEGER, yevmiye NUMERIC, tarih TEXT)"))
            
            # Eski tablolara WBS Kodu sütunu ekleme (Eğer yoksa)
            s.execute(text("ALTER TABLE satin_alma ADD COLUMN IF NOT EXISTS wbs_kodu TEXT"))
            s.execute(text("ALTER TABLE iscilik ADD COLUMN IF NOT EXISTS wbs_kodu TEXT"))

            # 2. YENİ TABLOLAR (CPM, Metraj, Tarihçe)
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
# 3. MÜHENDİSLİK MOTORLARI (CPM & EVM)
# ============================================================
def calculate_cpm(act_df, dep_df):
    """Gelişmiş CPM: FS, SS, FF destekli ve Döngü Korumalı."""
    if act_df.empty: return pd.DataFrame()
    
    G = nx.DiGraph()
    for _, row in act_df.iterrows(): G.add_node(row['wbs_code'])
    for _, row in dep_df.iterrows(): G.add_edge(row['predecessor'], row['successor'])

    if not nx.is_directed_acyclic_graph(G):
        st.error("⚠️ HATA: Döngüsel bağımlılık (A->B->C->A) tespit edildi! Lütfen İş Programı sekmesinden bağımlılıkları düzeltin.")
        return pd.DataFrame()

    topo = list(nx.topological_sort(G))
    es, ef, ls, lf = {}, {}, {}, {}

    # İleri Geçiş (Forward Pass)
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

    # Geri Geçiş (Backward Pass)
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
        
        # Karma İlerleme Hesabı (%30 Manuel, %70 Metraj)
        qty_prog = (row['act_qty'] / row['plan_qty'] * 100) if row['plan_qty'] > 0 else 0
        final_prog = (row['manual_progress'] * 0.3) + (qty_prog * 0.7)
        
        res.append({
            "WBS": wbs, "İş Kalemi": row['task_name'], "Süre": row['duration'],
            "Bütçe": row['budget'], "ES": es[wbs], "EF": ef[wbs],
            "Bolluk": slack, "Kritik": "EVET" if slack == 0 else "HAYIR",
            "Fiziksel İlerleme (%)": round(final_prog, 1)
        })
    return pd.DataFrame(res)

def calculate_evm(cpm_df, satin_df, iscilik_df):
    """Gerçek Maliyet (AC) hesaplamalı EVM Motoru"""
    if cpm_df.empty: return {}
    
    bac = cpm_df['Bütçe'].sum()
    
    # EV (Kazanılmış Değer) = Bütçe * Fiziksel İlerleme
    ev = sum((row['Bütçe'] * row['Fiziksel İlerleme (%)'] / 100) for _, row in cpm_df.iterrows())
    
    # AC (Gerçekleşen Maliyet) = Satın Alma + İşçilik
    ac_sat = satin_df['tutar'].sum() if not satin_df.empty else 0
    ac_isc = (iscilik_df['adet'] * iscilik_df['yevmiye']).sum() if not iscilik_df.empty else 0
    ac = ac_sat + ac_isc
    
    # PV (Planlanan Değer) - Zaman bazlı (Basit simülasyon: Projenin %50'sindeyiz)
    pv = bac * 0.5 
    
    cpi = ev / ac if ac > 0 else 0
    spi = ev / pv if pv > 0 else 0
    eac = bac / cpi if cpi > 0 else bac
    
    return {"BAC": bac, "PV": pv, "EV": ev, "AC": ac, "CPI": round(cpi,2), "SPI": round(spi,2), "EAC": eac}

# ============================================================
# 4. ANA UYGULAMA VE YAN MENÜ
# ============================================================
st.markdown('<div class="header-container"><h1 class="header-title">🏗️ EMİR ASEL İNŞAAT — PROJE YÖNETİM PANELİ</h1></div>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 📌 Navigasyon")
    menu = st.radio("Sayfa Seçin", [
        "📊 Dashboard", 
        "📅 İş Programı & CPM", 
        "📈 İlerleme & Metraj", 
        "🛒 Satın Alma", 
        "👷 İşçilik & Puantaj",
        "📑 Raporlama"
    ])

# Verileri Çek
df_act = conn.query("SELECT * FROM activities", ttl=0)
df_dep = conn.query("SELECT * FROM dependencies", ttl=0)
df_sat = conn.query("SELECT * FROM satin_alma", ttl=0)
df_isc = conn.query("SELECT * FROM iscilik", ttl=0)

cpm_df = calculate_cpm(df_act, df_dep) if not df_act.empty else pd.DataFrame()

# ============================================================
# SAYFA 1: DASHBOARD
# ============================================================
if menu == "📊 Dashboard":
    st.subheader("Yönetici Özeti")
    
    if df_act.empty:
        st.info("👋 Hoş geldiniz! Başlamak için lütfen 'İş Programı' sekmesinden WBS kalemlerini ekleyin.")
    else:
        evm = calculate_evm(cpm_df, df_sat, df_isc)
        avg_prog = cpm_df['Fiziksel İlerleme (%)'].mean()

        # KPI Kartları (Satır 1)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Toplam Bütçe (BAC)", f"{evm['BAC']:,.0f} ₺")
        c2.metric("Gerçekleşen Maliyet (AC)", f"{evm['AC']:,.0f} ₺", delta=f"{(evm['AC']-evm['EV']):,.0f} ₺", delta_color="inverse")
        c3.metric("Kazanılmış Değer (EV)", f"{evm['EV']:,.0f} ₺")
        c4.metric("Fiziksel İlerleme", f"%{avg_prog:.1f}")

        # KPI Kartları (Satır 2)
        c5, c6, c7, c8 = st.columns(4)
        c5.metric("Maliyet Performansı (CPI)", evm['CPI'], help="<1 ise bütçe aşımı var.")
        c6.metric("Zaman Performansı (SPI)", evm['SPI'], help="<1 ise proje geride.")
        c7.metric("Tahmini Bitiş Maliyeti (EAC)", f"{evm['EAC']:,.0f} ₺")
        c8.metric("Kalan Maliyet (ETC)", f"{evm['EAC'] - evm['AC']:,.0f} ₺")

        st.markdown("---")
        
        # Grafikler (S-Curve ve Maliyet Dağılımı)
        col_graph1, col_graph2 = st.columns([2, 1])
        with col_graph1:
            st.markdown("#### 📈 S-Eğrisi (Planlanan vs Gerçekleşen)")
            days = np.arange(0, 100, 10)
            plan_curve = 1 / (1 + np.exp(-0.1 * (days - 50))) * 100
            actual_curve = plan_curve * (avg_prog / 50)
            
            fig_scurve = go.Figure()
            fig_scurve.add_trace(go.Scatter(x=days, y=plan_curve, mode='lines', name='Planlanan (PV)', line=dict(color='#0A192F', width=3)))
            fig_scurve.add_trace(go.Scatter(x=days, y=actual_curve, mode='lines', name='Gerçekleşen (EV)', line=dict(color='#E63946', width=3, dash='dash')))
            fig_scurve.update_layout(xaxis_title="Proje Süresi (%)", yaxis_title="İlerleme (%)", margin=dict(l=0, r=0, t=10, b=0))
            st.plotly_chart(fig_scurve, use_container_width=True)

        with col_graph2:
            st.markdown("#### 💰 Maliyet Dağılımı")
            ac_sat = df_sat['tutar'].sum() if not df_sat.empty else 0
            ac_isc = (df_isc['adet'] * df_isc['yevmiye']).sum() if not df_isc.empty else 0
            fig_pie = px.pie(values=[ac_sat, ac_isc], names=["Satın Alma", "İşçilik"], 
                         color_discrete_sequence=['#E63946', '#0A192F'], hole=0.5)
            fig_pie.update_layout(margin=dict(l=0, r=0, t=10, b=0), showlegend=False)
            st.plotly_chart(fig_pie, use_container_width=True)

# ============================================================
# SAYFA 2: İŞ PROGRAMI & CPM
# ============================================================
elif menu == "📅 İş Programı & CPM":
    st.subheader("WBS ve Kritik Yol (CPM)")
    
    tab1, tab2, tab3 = st.tabs(["1. Aktiviteleri Yönet", "2. Bağımlılıkları Yönet", "3. Gantt & Kritik Yol"])
    
    with tab1:
        with st.form("add_activity"):
            st.markdown("##### Yeni Aktivite Ekle / Güncelle")
            c1, c2, c3, c4 = st.columns(4)
            wbs = c1.text_input("WBS Kodu (Örn: A01)")
            name = c2.text_input("İş Kalemi Adı")
            dur = c3.number_input("Süre (Gün)", min_value=1, value=1)
            bud = c4.number_input("Bütçe (TL)", min_value=0.0, value=0.0)
            if st.form_submit_button("Kaydet"):
                with conn.session as s:
                    s.execute(text("""
                        INSERT INTO activities (wbs_code, task_name, duration, budget)
                        VALUES (:w, :n, :d, :b)
                        ON CONFLICT (wbs_code) DO UPDATE SET task_name=:n, duration=:d, budget=:b
                    """), {"w": wbs, "n": name, "d": dur, "b": bud})
                    s.commit()
                st.rerun()
        st.dataframe(df_act[['wbs_code', 'task_name', 'duration', 'budget']], use_container_width=True)

    with tab2:
        with st.form("add_dep"):
            st.markdown("##### Yeni Bağımlılık Ekle")
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
            # Gantt Şeması
            start_date = datetime(2026, 1, 15)
            gantt_data = []
            for _, row in cpm_df.iterrows():
                gantt_data.append(dict(
                    Task=row['İş Kalemi'],
                    Start=start_date + timedelta(days=int(row['ES'])-1),
                    Finish=start_date + timedelta(days=int(row['EF'])-1),
                    Durum="Kritik" if row['Kritik'] == "EVET" else "Normal"
                ))
            
            fig_gantt = px.timeline(pd.DataFrame(gantt_data), x_start="Start", x_end="Finish", y="Task", color="Durum",
                              color_discrete_map={"Kritik": "#E63946", "Normal": "#0A192F"})
            fig_gantt.update_yaxes(autorange="reversed")
            st.plotly_chart(fig_gantt, use_container_width=True)
            
            st.markdown("#### CPM Tablosu")
            st.dataframe(cpm_df.style.applymap(lambda x: 'background-color: #ffcccc' if x == "EVET" else '', subset=['Kritik']), use_container_width=True)

# ============================================================
# SAYFA 3: İLERLEME & METRAJ
# ============================================================
elif menu == "📈 İlerleme & Metraj":
    st.subheader("Karma İlerleme Girişi (Manuel + Metraj)")
    
    if df_act.empty:
        st.warning("Lütfen önce İş Programı sekmesinden aktivite ekleyin.")
    else:
        with st.form("progress_form"):
            st.markdown("#### İmalat ve İlerleme Güncellemesi")
            selected_wbs = st.selectbox("İş Kalemi Seçin", df_act['wbs_code'] + " - " + df_act['task_name'])
            wbs_code = selected_wbs.split(" - ")[0]
            
            act_data = df_act[df_act['wbs_code'] == wbs_code].iloc[0]
            
            c1, c2 = st.columns(2)
            plan_qty = c1.number_input("Planlanan Miktar", value=float(act_data['plan_qty']))
            act_qty = c2.number_input("Gerçekleşen Miktar", value=float(act_data['act_qty']))
            
            man_prog = st.slider("Manuel İlerleme (%)", 0, 100, int(act_data['manual_progress']))
            
            if st.form_submit_button("Kaydet ve Geçmişe İşle"):
                with conn.session as s:
                    s.execute(text("UPDATE activities SET plan_qty=:pq, act_qty=:aq, manual_progress=:mp WHERE wbs_code=:w"), 
                              {"pq": plan_qty, "aq": act_qty, "mp": man_prog, "w": wbs_code})
                    s.execute(text("INSERT INTO progress_history (wbs_code, progress, note) VALUES (:w, :p, :n)"),
                              {"w": wbs_code, "p": man_prog, "n": f"Metraj: {act_qty}/{plan_qty}"})
                    s.commit()
                st.success("İlerleme tarihçeye kaydedildi.")
                st.rerun()
                
        st.markdown("---")
        st.subheader("Son İlerleme Hareketleri (Tarihçe)")
        df_hist = conn.query("SELECT * FROM progress_history ORDER BY log_date DESC LIMIT 10", ttl=0)
        st.dataframe(df_hist, use_container_width=True)

# ============================================================
# SAYFA 4: SATIN ALMA
# ============================================================
elif menu == "🛒 Satın Alma":
    st.subheader("Satın Alma (Gerçek Maliyet Girişi)")
    
    with st.form("satin_alma_form"):
        c1, c2, c3, c4 = st.columns(4)
        wbs_sat = c1.selectbox("İlgili WBS Kodu", df_act['wbs_code'] if not df_act.empty else ["Yok"])
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
    st.dataframe(df_sat, use_container_width=True)

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
    st.dataframe(df_isc, use_container_width=True)

# ============================================================
# SAYFA 6: RAPORLAMA (PDF & EXCEL)
# ============================================================
elif menu == "📑 Raporlama":
    st.subheader("Yönetici Raporları")
    
    if not df_act.empty:
        evm = calculate_evm(cpm_df, df_sat, df_isc)
        avg_prog = cpm_df['Fiziksel İlerleme (%)'].mean()
        
        # Excel İndirme
        st.markdown("### 📊 Excel Verileri")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_act.to_excel(writer, sheet_name='Aktiviteler', index=False)
            df_sat.to_excel(writer, sheet_name='Satın Alma', index=False)
            df_isc.to_excel(writer, sheet_name='İşçilik', index=False)
        st.download_button("📥 Tüm Verileri Excel Olarak İndir", data=output.getvalue(), file_name="EmirAsel_Veriler.xlsx")

        # PDF İndirme
        st.markdown("### 📄 Kurumsal Yönetici Raporu")
        if st.button("PDF Raporu Oluştur"):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            styles = getSampleStyleSheet()
            elements = []
            
            title_style = ParagraphStyle('Title', parent=styles['Heading1'], textColor=colors.HexColor('#E63946'))
            elements.append(Paragraph("EMİR ASEL İNŞAAT - YÖNETİCİ ÖZET RAPORU", title_style))
            elements.append(Spacer(1, 12))
            
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
            
            st.download_button("📥 PDF Raporunu İndir", data=buffer.getvalue(), file_name="EmirAsel_Rapor.pdf", mime="application/pdf")
