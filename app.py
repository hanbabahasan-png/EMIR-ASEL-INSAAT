import io
import re
from datetime import date, datetime, timedelta

import networkx as nx
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from sqlalchemy import text

# ============================================================
# EMİR ASEL İNŞAAT | Tek proje yönetim uygulaması
# Üretim Sertleştirilmiş Sürüm
# ============================================================

st.set_page_config(
    page_title="Emir Asel İnşaat | Proje Yönetimi",
    page_icon="🏗️",
    layout="wide",
)

st.markdown("""
<style>
.stApp {background:#F4F7F6}
.header-container {
 background:linear-gradient(135deg,#0A192F,#183153);
 padding:24px 28px;border-radius:14px;color:white;margin-bottom:20px
}
.header-title {color:white!important;font-size:27px!important;font-weight:800!important;margin:0}
div[data-testid="stMetric"] {
 background:white;border-radius:11px;padding:13px 16px;
 box-shadow:0 2px 8px rgba(0,0,0,.045);border-left:5px solid #B42335
}
section[data-testid="stSidebar"] {background:#fff;border-right:1px solid #e8e8e8}
</style>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="header-container"><h1 class="header-title">'
    '🏗️ EMİR ASEL İNŞAAT — PROJE YÖNETİM PANELİ</h1></div>',
    unsafe_allow_html=True,
)

# ------------------------- DB -------------------------
@st.cache_resource
def get_connection():
    try:
        return st.connection("neon", type="sql")
    except Exception as exc:
        st.error("Neon bağlantısı kurulamadı. .streamlit/secrets.toml ayarlarını kontrol edin.")
        st.exception(exc)
        st.stop()

conn = get_connection()

def run_sql(sql, params=None, fetch=False):
    with conn.session as session:
        result = session.execute(text(sql), params or {})
        if fetch:
            rows = result.fetchall()
            return pd.DataFrame(rows, columns=result.keys())
        session.commit()
        # Veri yazma işleminden sonra cache'i temizle
        st.cache_data.clear()
        return None

def init_db():
    statements = [
        """CREATE TABLE IF NOT EXISTS ilerleme (wbs_kodu TEXT PRIMARY KEY, ilerleme_yuzdesi INTEGER DEFAULT 0)""",
        """CREATE TABLE IF NOT EXISTS satin_alma (id SERIAL PRIMARY KEY, malzeme TEXT, tedarikci TEXT, tutar NUMERIC, durum TEXT, tarih TEXT)""",
        """CREATE TABLE IF NOT EXISTS iscilik (id SERIAL PRIMARY KEY, ekip TEXT, is_tanimi TEXT, adet INTEGER, yevmiye NUMERIC, tarih TEXT)""",
        "ALTER TABLE satin_alma ADD COLUMN IF NOT EXISTS wbs_kodu TEXT",
        "ALTER TABLE iscilik ADD COLUMN IF NOT EXISTS wbs_kodu TEXT",
        """CREATE TABLE IF NOT EXISTS activities (wbs_code TEXT PRIMARY KEY, task_name TEXT NOT NULL, duration INTEGER NOT NULL DEFAULT 1, budget NUMERIC NOT NULL DEFAULT 0, responsible TEXT, plan_qty NUMERIC NOT NULL DEFAULT 0, act_qty NUMERIC NOT NULL DEFAULT 0, unit TEXT, manual_progress NUMERIC NOT NULL DEFAULT 0, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        "ALTER TABLE activities ADD COLUMN IF NOT EXISTS responsible TEXT",
        "ALTER TABLE activities ADD COLUMN IF NOT EXISTS plan_qty NUMERIC DEFAULT 0",
        "ALTER TABLE activities ADD COLUMN IF NOT EXISTS act_qty NUMERIC DEFAULT 0",
        "ALTER TABLE activities ADD COLUMN IF NOT EXISTS unit TEXT",
        "ALTER TABLE activities ADD COLUMN IF NOT EXISTS manual_progress NUMERIC DEFAULT 0",
        "ALTER TABLE activities ADD COLUMN IF NOT EXISTS created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        "ALTER TABLE activities ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        """CREATE TABLE IF NOT EXISTS dependencies (id SERIAL PRIMARY KEY, predecessor TEXT NOT NULL, successor TEXT NOT NULL, relation_type TEXT NOT NULL DEFAULT 'FS')""",
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_dep_unique ON dependencies(predecessor, successor)",
        """CREATE TABLE IF NOT EXISTS progress_history (id SERIAL PRIMARY KEY, wbs_code TEXT, log_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP, progress NUMERIC, note TEXT)""",
        """CREATE TABLE IF NOT EXISTS audit_log (id SERIAL PRIMARY KEY, log_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP, username TEXT, action TEXT, details TEXT)""",
        "ALTER TABLE audit_log ADD COLUMN IF NOT EXISTS username TEXT",
        """CREATE TABLE IF NOT EXISTS cashflow (id SERIAL PRIMARY KEY, cash_date DATE, income NUMERIC DEFAULT 0, expense NUMERIC DEFAULT 0, forecast BOOLEAN DEFAULT TRUE)""",
        """CREATE TABLE IF NOT EXISTS stock (id SERIAL PRIMARY KEY, wbs_code TEXT, material TEXT, quantity NUMERIC DEFAULT 0, unit TEXT)""",
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_stock_unique ON stock(wbs_code, material)",
        """CREATE TABLE IF NOT EXISTS stock_movements (id SERIAL PRIMARY KEY, stock_id INTEGER REFERENCES stock(id) ON DELETE CASCADE, movement_type TEXT, quantity NUMERIC, movement_date DATE)""",
        """CREATE TABLE IF NOT EXISTS risk_register (id SERIAL PRIMARY KEY, description TEXT, impact TEXT, probability TEXT, mitigation TEXT)""",
        """CREATE TABLE IF NOT EXISTS documents (id SERIAL PRIMARY KEY, wbs_code TEXT, file_name TEXT, file_data BYTEA, upload_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""",
        "ALTER TABLE documents ADD COLUMN IF NOT EXISTS file_data BYTEA",
        """CREATE TABLE IF NOT EXISTS progress_payment (id SERIAL PRIMARY KEY, wbs_code TEXT, amount NUMERIC, payment_date DATE, status TEXT)""",
        """CREATE TABLE IF NOT EXISTS project_settings (setting_key TEXT PRIMARY KEY, setting_value TEXT)""",
        """CREATE TABLE IF NOT EXISTS overhead_costs (id SERIAL PRIMARY KEY, category TEXT, description TEXT, amount NUMERIC, payment_date DATE, status TEXT)"""
    ]
    with conn.session as session:
        for statement in statements:
            session.execute(text(statement))
        session.commit()

try:
    init_db()
except Exception as exc:
    st.error("Veritabanı şeması hazırlanamadı.")
    st.exception(exc)
    st.stop()

def log_action(action, details):
    try:
        user = "Yönetici"
        run_sql("INSERT INTO audit_log(username, action, details) VALUES (:u, :a, :d)",
                {"u": user, "a": str(action)[:200], "d": str(details)[:4000]})
    except Exception:
        pass

@st.cache_data(ttl=30)
def load_df(table, order=None):
    allowed = {
        "activities", "dependencies", "satin_alma", "iscilik",
        "progress_history", "cashflow", "stock", "risk_register",
        "documents", "progress_payment", "audit_log", "overhead_costs"
    }
    if table not in allowed:
        raise ValueError("Geçersiz tablo")
    suffix = f" ORDER BY {order}" if order else ""
    return run_sql(f"SELECT * FROM {table}{suffix}", fetch=True)

def valid_wbs(value):
    return bool(re.fullmatch(r"[A-Z][0-9]{2}", str(value).strip().upper()))

def number(value, default=0.0):
    try:
        n = float(value)
        return n if np.isfinite(n) else default
    except (TypeError, ValueError):
        return default

def get_project_start():
    df = run_sql("SELECT setting_value FROM project_settings WHERE setting_key='start_date'", fetch=True)
    if not df.empty:
        try:
            return date.fromisoformat(df.iloc[0, 0])
        except Exception:
            pass
    return date(2026, 1, 15)

def save_project_start(start):
    run_sql("""
        INSERT INTO project_settings(setting_key,setting_value)
        VALUES ('start_date',:v)
        ON CONFLICT(setting_key) DO UPDATE SET setting_value=EXCLUDED.setting_value
    """, {"v": start.isoformat()})
    log_action("PROJE BAŞLANGIÇ TARİHİ", start.isoformat())

# --------------------- Progress & CPM ---------------------
def activity_progress(row):
    plan = number(row.get("plan_qty"))
    actual = number(row.get("act_qty"))
    manual = min(100.0, max(0.0, number(row.get("manual_progress"))))
    quantity_pct = min(100.0, max(0.0, actual / plan * 100)) if plan > 0 else manual
    return min(100.0, max(0.0, 0.7 * quantity_pct + 0.3 * manual)) if plan > 0 else manual

def calculate_cpm(activities, dependencies):
    if activities.empty:
        return pd.DataFrame(), None
    acts = activities.copy()
    acts["wbs_code"] = acts["wbs_code"].astype(str).str.upper()
    act_map = acts.set_index("wbs_code").to_dict("index")
    graph = nx.DiGraph()
    graph.add_nodes_from(act_map)

    valid_deps = []
    if not dependencies.empty:
        for _, dep in dependencies.iterrows():
            p, s = str(dep["predecessor"]).upper(), str(dep["successor"]).upper()
            rel = str(dep.get("relation_type", "FS")).upper()
            if p not in act_map or s not in act_map or p == s:
                continue
            if rel not in {"FS", "SS", "FF", "SF"}:
                continue
            graph.add_edge(p, s)
            valid_deps.append((p, s, rel))

    if not nx.is_directed_acyclic_graph(graph):
        return pd.DataFrame(), "Bağımlılıklarda döngü var. CPM hesaplanamadı."

    topo = list(nx.topological_sort(graph))
    es, ef, ls, lf = {}, {}, {}, {}
    pred_map = {n: [] for n in topo}
    succ_map = {n: [] for n in topo}
    for p, s, rel in valid_deps:
        pred_map[s].append((p, rel))
        succ_map[p].append((s, rel))

    for node in topo:
        dur = max(1, int(number(act_map[node].get("duration"), 1)))
        starts = [1]
        for p, rel in pred_map[node]:
            pdur = max(1, int(number(act_map[p].get("duration"), 1)))
            if rel == "FS": starts.append(ef[p] + 1)
            elif rel == "SS": starts.append(es[p])
            elif rel == "FF": starts.append(ef[p] - dur + 1)
            elif rel == "SF": starts.append(es[p] - dur + 1)
        es[node] = max(starts)
        ef[node] = es[node] + dur - 1

    project_finish = max(ef.values())
    for node in reversed(topo):
        dur = max(1, int(number(act_map[node].get("duration"), 1)))
        if not succ_map[node]:
            lf[node] = project_finish
        else:
            limits = []
            for s, rel in succ_map[node]:
                sdur = max(1, int(number(act_map[s].get("duration"), 1)))
                if rel == "FS": limits.append(ls[s] - 1)
                elif rel == "SS": limits.append(ls[s] + dur - 1)
                elif rel == "FF": limits.append(lf[s])
                elif rel == "SF": limits.append(lf[s] + dur - 1)
            lf[node] = min(limits)
        ls[node] = lf[node] - dur + 1

    rows = []
    for node in topo:
        a = act_map[node]
        slack = ls[node] - es[node]
        rows.append({
            "WBS": node, "İş Kalemi": a.get("task_name", ""),
            "Süre (Gün)": int(number(a.get("duration"), 1)),
            "Bütçe (TL)": number(a.get("budget")),
            "ES": es[node], "EF": ef[node], "LS": ls[node], "LF": lf[node],
            "Bolluk (Gün)": slack,
            "Kritik": "EVET" if slack == 0 else "HAYIR",
            "Fiziksel İlerleme (%)": round(activity_progress(a), 1),
        })
    return pd.DataFrame(rows), None

def calculate_evm(cpm, purchases, labor, overheads, start_date):
    if cpm.empty:
        return {}
    bac = float(cpm["Bütçe (TL)"].sum())
    ev = float((cpm["Bütçe (TL)"] * cpm["Fiziksel İlerleme (%)"] / 100).sum())

    paid_sat = purchases[purchases["durum"].astype(str).str.lower().isin(["ödendi", "odendi"])] if not purchases.empty else purchases
    ac_purchase = float(pd.to_numeric(paid_sat["tutar"], errors="coerce").fillna(0).sum()) if not paid_sat.empty else 0.0
    
    ac_labor = float((pd.to_numeric(labor["adet"], errors="coerce").fillna(0) * pd.to_numeric(labor["yevmiye"], errors="coerce").fillna(0)).sum()) if not labor.empty else 0.0
    
    paid_overhead = overheads[overheads["status"].astype(str).str.lower().isin(["ödendi", "odendi"])] if not overheads.empty else overheads
    ac_overhead = float(pd.to_numeric(paid_overhead["amount"], errors="coerce").fillna(0).sum()) if not paid_overhead.empty else 0.0

    ac = ac_purchase + ac_labor + ac_overhead

    elapsed = max(0, (date.today() - start_date).days + 1)
    pv = 0.0
    for _, row in cpm.iterrows():
        es, ef = int(row["ES"]), int(row["EF"])
        duration = max(1, int(row["Süre (Gün)"]))
        budget = number(row["Bütçe (TL)"])
        if elapsed >= ef:
            pv += budget
        elif elapsed >= es:
            pv += budget * (elapsed - es + 1) / duration

    cpi = ev / ac if ac > 0 else 0.0
    spi = ev / pv if pv > 0 else 0.0
    
    if cpi > 0:
        eac = ac + ((bac - ev) / cpi)
    else:
        eac = bac
        
    return {"BAC": bac, "PV": pv, "EV": ev, "AC": ac, "CPI": cpi, "SPI": spi, "EAC": eac, "ETC": eac - ac}

# ------------------------- Sidebar -------------------------
with st.sidebar:
    try:
        st.image(Image.open("logo.png"), use_container_width=True)
    except Exception:
        st.markdown("## EMİR ASEL")
    st.markdown("### 📌 Menü")
    menu = st.radio("Sayfa", [
        "📊 Dashboard", "📅 İş Programı & CPM", "📈 İlerleme & Metraj",
        "🛒 Satın Alma", "👷 İşçilik & Puantaj", "💼 Muhasebe & Cari", 
        "📊 Finansal Kontrol", "📦 Stok & Nakit Akışı", 
        "⚠️ Risk & Ödeme", "📁 Dokümanlar", 
        "⚙️ Veri Aktarımı & Raporlar", "🧾 İşlem Kayıtları",
    ])
    project_start = st.date_input("Proje başlangıç tarihi", value=get_project_start())
    if st.button("Başlangıç tarihini kaydet", use_container_width=True):
        save_project_start(project_start)
        st.success("Tarih kaydedildi.")

df_act = load_df("activities", "wbs_code")
df_dep = load_df("dependencies", "id")
df_sat = load_df("satin_alma", "id DESC")
df_isc = load_df("iscilik", "id DESC")
df_overhead = load_df("overhead_costs", "id DESC")
df_risk = load_df("risk_register", "id DESC")
df_stock = load_df("stock", "id DESC")
cpm_df, cpm_error = calculate_cpm(df_act, df_dep)

if cpm_error:
    st.error(cpm_error)

def wbs_options():
    return df_act["wbs_code"].astype(str).tolist() if not df_act.empty else []

# ------------------------- Dashboard -------------------------
if menu == "📊 Dashboard":
    st.subheader("Yönetici Özeti")
    if df_act.empty:
        st.info("Başlamak için İş Programı bölümünden aktivite ekleyin veya Excel aktarımı yapın.")
    elif not cpm_df.empty:
        evm = calculate_evm(cpm_df, df_sat, df_isc, df_overhead, project_start)
        avg_progress = float(cpm_df["Fiziksel İlerleme (%)"].mean())
        critical_count = int((cpm_df["Kritik"] == "EVET").sum())
        
        health = 100
        if evm["CPI"] and evm["CPI"] < 1: health -= 20
        if evm["SPI"] and evm["SPI"] < 1: health -= 20
        if critical_count > 10: health -= 10
        
        if not df_risk.empty:
            impact_map = {"Düşük": 1, "Orta": 2, "Yüksek": 3, "Kritik": 4}
            prob_map = {"Düşük": 1, "Orta": 2, "Yüksek": 3}
            
            df_risk_copy = df_risk.copy()
            df_risk_copy['impact_score'] = df_risk_copy['impact'].map(impact_map).fillna(1)
            df_risk_copy['prob_score'] = df_risk_copy['probability'].map(prob_map).fillna(1)
            total_risk_score = (df_risk_copy['impact_score'] * df_risk_copy['prob_score']).sum()
            risk_avg = total_risk_score / len(df_risk) if len(df_risk) > 0 else 0
            
            if risk_avg > 6: health -= 15
            elif risk_avg > 3: health -= 10
            
        if health < 0: health = 0

        cols = st.columns(5)
        metrics = [
            ("Toplam Bütçe (BAC)", f'{evm["BAC"]:,.0f} ₺'),
            ("Gerçekleşen Maliyet (AC)", f'{evm["AC"]:,.0f} ₺'),
            ("Kazanılmış Değer (EV)", f'{evm["EV"]:,.0f} ₺'),
            ("Fiziksel İlerleme", f"%{avg_progress:.1f}"),
            ("Proje Sağlığı", f"{health}/100"),
        ]
        for col, (label, value) in zip(cols, metrics):
            col.metric(label, value)
            
        cols2 = st.columns(4)
        for col, (label, value) in zip(cols2, [
            ("CPI", f'{evm["CPI"]:.2f}'), ("SPI", f'{evm["SPI"]:.2f}'),
            ("Tahmini Bitiş Maliyeti", f'{evm["EAC"]:,.0f} ₺'),
            ("Tahmini Kalan Maliyet", f'{evm["ETC"]:,.0f} ₺'),
        ]):
            col.metric(label, value)

        st.markdown("---")
        completed_act = len(cpm_df[cpm_df["Fiziksel İlerleme (%)"] >= 100])
        active_act = len(cpm_df[(cpm_df["Fiziksel İlerleme (%)"] > 0) & (cpm_df["Fiziksel İlerleme (%)"] < 100)])
        total_stock = len(df_stock)
        open_risks = len(df_risk)

        cols3 = st.columns(4)
        cols3[0].metric("Tamamlanan Aktivite", completed_act)
        cols3[1].metric("Aktif Aktivite", active_act)
        cols3[2].metric("Stok Kartı Sayısı", total_stock)
        cols3[3].metric("Açık Risk Sayısı", open_risks)

        st.markdown("#### Aktivite ilerleme durumu")
        chart = cpm_df[["WBS", "Fiziksel İlerleme (%)"]].sort_values("Fiziksel İlerleme (%)")
        fig = px.bar(chart, x="Fiziksel İlerleme (%)", y="WBS", orientation="h",
                     range_x=[0, 100], text="Fiziksel İlerleme (%)")
        st.plotly_chart(fig, use_container_width=True)

# ------------------------- Dokümanlar -------------------------
elif menu == "📁 Dokümanlar":
    st.subheader("Proje Doküman Yönetimi")
    st.info("Fatura, irsaliye, ruhsat, proje çizimleri veya fotoğrafları buradan WBS kodlarına bağlayarak sisteme kaydedebilir ve daha sonra indirebilirsiniz. (Maks 20 MB)")
    
    with st.form("doc_upload_form"):
        c1, c2 = st.columns(2)
        doc_wbs = c1.selectbox("İlgili WBS Kodu", ["GENEL"] + wbs_options())
        uploaded_file = c2.file_uploader("Dosya Seçin (PDF, Excel, Resim)", type=["pdf", "xlsx", "xls", "docx", "png", "jpg", "jpeg"])
        
        if st.form_submit_button("Dokümanı Sisteme Kaydet"):
            if uploaded_file is not None:
                file_bytes = uploaded_file.getvalue()
                if len(file_bytes) > 20 * 1024 * 1024:
                    st.error("Dosya boyutu 20 MB sınırını aşıyor. Lütfen daha küçük bir dosya seçin.")
                else:
                    file_name = uploaded_file.name
                    run_sql("""INSERT INTO documents (wbs_code, file_name, file_data) VALUES (:w, :f, :d)""", 
                            {"w": doc_wbs, "f": file_name, "d": file_bytes})
                    log_action("DOKÜMAN YÜKLEME", f"{doc_wbs} - {file_name}")
                    st.success(f"'{file_name}' başarıyla arşivlendi!")
                    st.rerun()
            else:
                st.error("Lütfen bir dosya seçin.")

    st.markdown("---")
    st.markdown("### 📂 Yüklenen Dokümanlar")
    df_docs = run_sql("SELECT id, wbs_code, file_name, upload_date FROM documents ORDER BY upload_date DESC", fetch=True)
    
    if not df_docs.empty:
        st.dataframe(df_docs, use_container_width=True, hide_index=True)
        
        st.markdown("#### 📥 Doküman İndir")
        doc_options = {f"{r.file_name} ({r.wbs_code})": r.id for r in df_docs.itertuples()}
        selected_doc = st.selectbox("İndirilecek dosyayı seçin", list(doc_options.keys()))
        doc_id = doc_options[selected_doc]
        
        file_data_df = run_sql("SELECT file_name, file_data FROM documents WHERE id=:id", {"id": doc_id}, fetch=True)
        if not file_data_df.empty:
            f_name = file_data_df.iloc[0]['file_name']
            f_data = file_data_df.iloc[0]['file_data']
            
            if isinstance(f_data, memoryview):
                f_data = f_data.tobytes()
            elif f_data is None:
                f_data = b""
                
            if f_data:
                st.download_button(
                    label="Dosyayı İndir", 
                    data=bytes(f_data), 
                    file_name=f_name
                )
    else:
        st.info("Henüz yüklenmiş bir doküman yok.")

# ------------------------- Finansal Kontrol -------------------------
elif menu == "📊 Finansal Kontrol":
    st.subheader("Bütçe vs Gerçekleşen Maliyet Karşılaştırması")
    if df_act.empty:
        st.info("Aktivite verisi yok.")
    else:
        df_temp = df_act.copy()
        df_temp['Kategori'] = df_temp['wbs_code'].astype(str).str[0]
        plan_budget = df_temp.groupby('Kategori')['budget'].sum().reset_index()
        plan_budget.columns = ['Kategori', 'Planlanan Bütçe (TL)']
        
        actual_sat = df_sat.copy()
        if not actual_sat.empty:
            actual_sat = actual_sat[actual_sat['durum'].astype(str).str.lower().isin(['ödendi', 'odendi'])]
            actual_sat['Kategori'] = actual_sat['wbs_kodu'].astype(str).str[0]
            sat_group = actual_sat.groupby('Kategori')['tutar'].sum().reset_index()
            sat_group.columns = ['Kategori', 'Gerçekleşen Satın Alma']
        else:
            sat_group = pd.DataFrame(columns=['Kategori', 'Gerçekleşen Satın Alma'])
            
        actual_isc = df_isc.copy()
        if not actual_isc.empty:
            actual_isc['Kategori'] = actual_isc['wbs_kodu'].astype(str).str[0]
            isc_group = actual_isc.groupby('Kategori').apply(lambda x: (x['adet'] * x['yevmiye']).sum(), include_groups=False).reset_index()
            isc_group.columns = ['Kategori', 'Gerçekleşen İşçilik']
        else:
            isc_group = pd.DataFrame(columns=['Kategori', 'Gerçekleşen İşçilik'])
            
        merged = pd.merge(plan_budget, sat_group, on='Kategori', how='outer').fillna(0)
        merged = pd.merge(merged, isc_group, on='Kategori', how='outer').fillna(0)
        merged['Toplam Harcama'] = merged['Gerçekleşen Satın Alma'] + merged['Gerçekleşen İşçilik']
        merged['Kalan Bütçe'] = merged['Planlanan Bütçe (TL)'] - merged['Toplam Harcama']
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=merged['Kategori'], y=merged['Planlanan Bütçe (TL)'], name='Planlanan', marker_color='#0A192F'))
        fig.add_trace(go.Bar(x=merged['Kategori'], y=merged['Toplam Harcama'], name='Gerçekleşen', marker_color='#E63946'))
        fig.update_layout(barmode='group', title="Kategori Bazlı Bütçe vs Harcama")
        st.plotly_chart(fig, use_container_width=True)
        
        st.dataframe(merged.style.format("{:,.0f}"), use_container_width=True)

        st.markdown("#### Proje Genel Giderleri (Muhasebe)")
        if not df_overhead.empty:
            total_oh = df_overhead[df_overhead['status'].astype(str).str.lower().isin(['ödendi', 'odendi'])]['amount'].sum()
            st.metric("Ödenen Genel Giderler (Arsa, Ruhsat vb.)", f"{total_oh:,.0f} ₺")
            st.dataframe(df_overhead[['category', 'description', 'amount', 'status']], use_container_width=True)
        else:
            st.info("Henüz muhasebe kaydı girilmemiş.")

# ------------------------- Muhasebe & Cari -------------------------
elif menu == "💼 Muhasebe & Cari":
    st.subheader("Genel Giderler ve Cari Hesaplar")
    tab1, tab2 = st.tabs(["Proje Genel Giderleri", "Tedarikçi Cari Hesapları"])
    
    with tab1:
        with st.form("overhead_form"):
            c = st.columns(3)
            category = c[0].selectbox("Kategori", ["Arsa", "Ruhsat/Harç", "Mimarlık/Proje", "Yapı Denetim", "Vergi", "Diğer"])
            desc = c[1].text_input("Açıklama")
            amount = c[2].number_input("Tutar (TL)", min_value=0.0)
            
            c2 = st.columns(2)
            p_date = c2[0].date_input("Tarih", value=date.today())
            status = c2[1].selectbox("Ödeme Durumu", ["Bekliyor", "Ödendi"])
            
            if st.form_submit_button("Genel Gider Kaydet"):
                if not desc.strip():
                    st.error("Açıklama zorunludur.")
                else:
                    run_sql("""INSERT INTO overhead_costs (category, description, amount, payment_date, status) 
                               VALUES (:c, :d, :a, :p, :s)""",
                            {"c": category, "d": desc, "a": amount, "p": p_date, "s": status})
                    log_action("GENEL GİDER", f"{category} - {amount} TL")
                    st.rerun()
        
        st.dataframe(load_df("overhead_costs", "id DESC"), use_container_width=True, hide_index=True)

    with tab2:
        if not df_sat.empty:
            df_sat_cari = df_sat.copy()
            df_sat_cari['tedarikci'] = df_sat_cari['tedarikci'].fillna('Bilinmeyen')
            
            cari = df_sat_cari.groupby('tedarikci').agg(
                Toplam_Borc=('tutar', 'sum'),
                Odenen=('tutar', lambda x: df_sat_cari.loc[x.index][df_sat_cari.loc[x.index, 'durum'].astype(str).str.lower().isin(['ödendi', 'odendi'])]['tutar'].sum())
            ).reset_index()
            cari['Kalan_Borc'] = cari['Toplam_Borc'] - cari['Odenen']
            
            st.dataframe(cari.style.format({
                'Toplam_Borc': '{:,.0f} ₺',
                'Odenen': '{:,.0f} ₺',
                'Kalan_Borc': '{:,.0f} ₺'
            }), use_container_width=True)
            
            fig = px.bar(cari, x='tedarikci', y=['Toplam_Borc', 'Odenen'], barmode='group', 
                         title="Tedarikçi Bazlı Borç ve Ödemeler", 
                         color_discrete_sequence=['#E63946', '#2A9D8F'])
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Henüz satın alma kaydı yok.")

# ------------------------- Schedule -------------------------
elif menu == "📅 İş Programı & CPM":
    st.subheader("WBS, Aktivite ve Kritik Yol Yönetimi")
    tab1, tab2, tab3 = st.tabs(["Aktiviteler", "Bağımlılıklar", "Gantt / CPM"])
    with tab1:
        with st.form("activity_form", clear_on_submit=True):
            st.markdown("##### Aktivite ekle / güncelle")
            c = st.columns(4)
            wbs = c[0].text_input("WBS (A01)")
            name = c[1].text_input("İş kalemi")
            duration = c[2].number_input("Süre (gün)", 1, 10000, 1)
            budget = c[3].number_input("Bütçe (TL)", 0.0, step=1000.0)
            c2 = st.columns(3)
            responsible = c2[0].text_input("Sorumlu")
            plan_qty = c2[1].number_input("Plan miktarı", min_value=0.0)
            unit = c2[2].text_input("Birim (m², m³, adet...)")
            if st.form_submit_button("Kaydet / Güncelle"):
                wbs = wbs.strip().upper()
                if not valid_wbs(wbs) or not name.strip():
                    st.error("WBS A01 biçiminde olmalı ve iş kalemi boş bırakılamaz.")
                else:
                    run_sql("""
                        INSERT INTO activities(wbs_code,task_name,duration,budget,responsible,plan_qty,unit)
                        VALUES(:w,:n,:d,:b,:r,:q,:u)
                        ON CONFLICT(wbs_code) DO UPDATE SET
                        task_name=EXCLUDED.task_name,duration=EXCLUDED.duration,
                        budget=EXCLUDED.budget,responsible=EXCLUDED.responsible,
                        plan_qty=EXCLUDED.plan_qty,unit=EXCLUDED.unit,
                        updated_at=CURRENT_TIMESTAMP
                    """, {"w": wbs, "n": name.strip(), "d": duration, "b": budget,
                          "r": responsible.strip(), "q": plan_qty, "u": unit.strip()})
                    log_action("AKTİVİTE KAYDEDİLDİ", f"{wbs} - {name}")
                    st.success("Aktivite kaydedildi.")
                    st.rerun()
        if not df_act.empty:
            st.dataframe(df_act, use_container_width=True, hide_index=True)
    with tab2:
        if len(wbs_options()) >= 2:
            with st.form("dependency_form"):
                c = st.columns(3)
                pred = c[0].selectbox("Öncül aktivite", wbs_options())
                succ_options = [x for x in wbs_options() if x != pred]
                succ = c[1].selectbox("Ardıl aktivite", succ_options)
                relation = c[2].selectbox("İlişki", ["FS", "SS", "FF", "SF"])
                if st.form_submit_button("Bağımlılığı kaydet"):
                    if pred == succ:
                        st.error("Bir aktivite kendisine bağlanamaz.")
                    else:
                        try:
                            run_sql("""INSERT INTO dependencies(predecessor,successor,relation_type)
                                       VALUES(:p,:s,:r) ON CONFLICT(predecessor,successor)
                                       DO UPDATE SET relation_type=EXCLUDED.relation_type""",
                                    {"p": pred, "s": succ, "r": relation})
                            test_dep = load_df("dependencies")
                            _, err = calculate_cpm(df_act, test_dep)
                            if err:
                                run_sql("DELETE FROM dependencies WHERE predecessor=:p AND successor=:s",
                                        {"p": pred, "s": succ})
                                st.error(err)
                            else:
                                log_action("BAĞIMLILIK", f"{pred} -> {succ} ({relation})")
                                st.success("Bağımlılık kaydedildi.")
                                st.rerun()
                        except Exception as exc:
                            st.error(f"Bağımlılık kaydedilemedi: {exc}")
        else:
            st.info("Bağımlılık tanımlamak için en az iki aktivite gerekir.")
        dep_view = load_df("dependencies", "id")
        st.dataframe(dep_view, use_container_width=True, hide_index=True)
        if not dep_view.empty:
            dep_id = st.selectbox("Silinecek bağımlılık ID", dep_view["id"].tolist())
            if st.button("Bağımlılığı sil"):
                run_sql("DELETE FROM dependencies WHERE id=:id", {"id": int(dep_id)})
                log_action("BAĞIMLILIK SİLİNDİ", dep_id)
                st.rerun()
    with tab3:
        if not cpm_df.empty:
            gantt = []
            for _, row in cpm_df.iterrows():
                gantt.append({
                    "Aktivite": f'{row["WBS"]} - {row["İş Kalemi"]}',
                    "Başlangıç": project_start + timedelta(days=int(row["ES"]) - 1),
                    "Bitiş": project_start + timedelta(days=int(row["EF"])), 
                    "Durum": "Kritik" if row["Kritik"] == "EVET" else "Normal",
                    "İlerleme": row["Fiziksel İlerleme (%)"],
                })
            fig = px.timeline(pd.DataFrame(gantt), x_start="Başlangıç", x_end="Bitiş",
                              y="Aktivite", color="Durum",
                              color_discrete_map={"Kritik": "#B42335", "Normal": "#183153"},
                              hover_data=["İlerleme"])
            fig.update_yaxes(autorange="reversed")
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(cpm_df, use_container_width=True, hide_index=True)

# ------------------------- Progress -------------------------
elif menu == "📈 İlerleme & Metraj":
    st.subheader("Metraj ve Fiziksel İlerleme")
    if df_act.empty:
        st.info("Önce aktivite ekleyin.")
    else:
        options = {f'{r.wbs_code} — {r.task_name}': r.wbs_code for r in df_act.itertuples()}
        selected = st.selectbox("Aktivite", list(options))
        wbs = options[selected]
        row = df_act[df_act["wbs_code"] == wbs].iloc[0]
        with st.form("progress_form"):
            c = st.columns(3)
            planned = c[0].number_input("Planlanan miktar", min_value=0.0, value=number(row["plan_qty"]))
            actual = c[1].number_input("Gerçekleşen miktar", min_value=0.0, value=number(row["act_qty"]))
            manual = c[2].slider("Manuel ilerleme (%)", 0, 100, int(number(row["manual_progress"])))
            note = st.text_input("Açıklama")
            preview_qty = min(100, actual / planned * 100) if planned > 0 else manual
            preview_final = .7 * preview_qty + .3 * manual if planned > 0 else manual
            st.caption(f"Hesaplanan fiziksel ilerleme: %{preview_final:.1f}")
            if st.form_submit_button("İlerlemeyi kaydet"):
                run_sql("""UPDATE activities SET plan_qty=:p,act_qty=:a,manual_progress=:m,
                           updated_at=CURRENT_TIMESTAMP WHERE wbs_code=:w""",
                        {"p": planned, "a": actual, "m": manual, "w": wbs})
                run_sql("""INSERT INTO progress_history(wbs_code,progress,note)
                           VALUES(:w,:p,:n)""",
                        {"w": wbs, "p": preview_final, "n": note or f"Metraj {actual}/{planned}"})
                log_action("İLERLEME", f"{wbs}: %{preview_final:.1f}")
                st.success("İlerleme ve tarihçe kaydedildi.")
                st.rerun()
    st.markdown("#### Son ilerleme kayıtları")
    st.dataframe(load_df("progress_history", "log_date DESC"), use_container_width=True, hide_index=True)

# ------------------------- Purchases -------------------------
elif menu == "🛒 Satın Alma":
    st.subheader("Satın Alma ve Tedarik Takibi")
    with st.form("purchase_form", clear_on_submit=True):
        c = st.columns(5)
        wbs = c[0].selectbox("WBS", wbs_options() or ["GENEL"])
        material = c[1].text_input("Malzeme / hizmet")
        supplier = c[2].text_input("Tedarikçi")
        amount = c[3].number_input("Tutar (TL)", min_value=0.0, step=100.0)
        status = c[4].selectbox("Durum", ["Teklif", "Sipariş Verildi", "Teslim Edildi", "Ödendi"])
        if st.form_submit_button("Satın alma kaydı ekle"):
            if not material.strip():
                st.error("Malzeme adı zorunludur.")
            else:
                run_sql("""INSERT INTO satin_alma(wbs_kodu,malzeme,tedarikci,tutar,durum,tarih)
                           VALUES(:w,:m,:t,:a,:d,:dt)""",
                        {"w": wbs, "m": material.strip(), "t": supplier.strip(),
                         "a": amount, "d": status, "dt": date.today().isoformat()})
                log_action("SATIN ALMA", f"{material} / {amount} TL / {status}")
                st.success("Kayıt eklendi.")
                st.rerun()
    st.dataframe(df_sat, use_container_width=True, hide_index=True)
    if not df_sat.empty:
        purchase_id = st.selectbox("Durumu güncellenecek kayıt ID", df_sat["id"].tolist())
        new_status = st.selectbox("Yeni durum", ["Teklif", "Sipariş Verildi", "Teslim Edildi", "Ödendi"])
        if st.button("Satın alma durumunu güncelle"):
            run_sql("UPDATE satin_alma SET durum=:d WHERE id=:id",
                    {"d": new_status, "id": int(purchase_id)})
            log_action("SATIN ALMA DURUMU", f"ID {purchase_id}: {new_status}")
            st.rerun()

# ------------------------- Labor -------------------------
elif menu == "👷 İşçilik & Puantaj":
    st.subheader("Ekip, Puantaj ve İşçilik Maliyeti")
    with st.form("labor_form", clear_on_submit=True):
        c = st.columns(5)
        wbs = c[0].selectbox("WBS", wbs_options() or ["GENEL"])
        team = c[1].text_input("Ekip")
        description = c[2].text_input("İş tanımı")
        count = c[3].number_input("Kişi sayısı", min_value=1, value=1)
        wage = c[4].number_input("Kişi başı günlük yevmiye (TL)", min_value=0.0)
        days = st.number_input("Çalışılan gün", min_value=1, value=1)
        if st.form_submit_button("Puantaj kaydet"):
            if not team.strip():
                st.error("Ekip adı zorunludur.")
            else:
                run_sql("""INSERT INTO iscilik(wbs_kodu,ekip,is_tanimi,adet,yevmiye,tarih)
                           VALUES(:w,:e,:i,:a,:y,:t)""",
                        {"w": wbs, "e": team.strip(), "i": description.strip(),
                         "a": int(count * days), "y": wage, "t": date.today().isoformat()})
                log_action("İŞÇİLİK", f"{team}: {count} kişi x {days} gün")
                st.success("Puantaj kaydedildi.")
                st.rerun()
    st.caption("Toplam işçilik maliyeti = kişi-gün (adet) × kişi başı günlük yevmiye.")
    st.dataframe(df_isc, use_container_width=True, hide_index=True)

# ------------------------- Stock & cashflow -------------------------
elif menu == "📦 Stok & Nakit Akışı":
    st.subheader("Stok ve Nakit Akışı")
    tab_stock, tab_cash = st.tabs(["Stok", "Nakit Akışı"])
    with tab_stock:
        with st.form("stock_form", clear_on_submit=True):
            c = st.columns(4)
            wbs = c[0].selectbox("WBS", wbs_options() or ["GENEL"])
            material = c[1].text_input("Malzeme")
            qty = c[2].number_input("Miktar", min_value=0.0)
            unit = c[3].text_input("Birim")
            if st.form_submit_button("Stok kartı ekle"):
                if material.strip():
                    run_sql("""
                        INSERT INTO stock(wbs_code, material, quantity, unit) 
                        VALUES(:w, :m, :q, :u)
                        ON CONFLICT(wbs_code, material) DO UPDATE SET 
                        quantity = stock.quantity + EXCLUDED.quantity,
                        unit = EXCLUDED.unit
                    """, {"w": wbs, "m": material.strip(), "q": qty, "u": unit.strip()})
                    log_action("STOK KARTI", f"{material} (+{qty})")
                    st.success("Stok kartı eklendi/güncellendi.")
                    st.rerun()
                else:
                    st.error("Malzeme adı girin.")
        stock_df = load_df("stock", "id DESC")
        st.dataframe(stock_df, use_container_width=True, hide_index=True)
        if not stock_df.empty:
            stock_id = st.selectbox("Hareket yapılacak stok kartı", stock_df["id"].tolist())
            with st.form("stock_movement_form"):
                movement = st.selectbox("Hareket türü", ["Giriş", "Çıkış"])
                movement_qty = st.number_input("Hareket miktarı", min_value=0.01)
                if st.form_submit_button("Stok hareketini kaydet"):
                    current = number(stock_df.loc[stock_df["id"] == stock_id, "quantity"].iloc[0])
                    new_qty = current + movement_qty if movement == "Giriş" else current - movement_qty
                    if new_qty < 0:
                        st.error("Stok çıkışı mevcut miktarı aşamaz.")
                    else:
                        run_sql("UPDATE stock SET quantity=:q WHERE id=:id",
                                {"q": new_qty, "id": int(stock_id)})
                        run_sql("""INSERT INTO stock_movements(stock_id,movement_type,quantity,movement_date)
                                   VALUES(:id,:t,:q,:d)""",
                                {"id": int(stock_id), "t": movement, "q": movement_qty, "d": date.today()})
                        log_action("STOK HAREKETİ", f"{stock_id}: {movement} {movement_qty}")
                        st.rerun()
    with tab_cash:
        with st.form("cashflow_form", clear_on_submit=True):
            c = st.columns(4)
            cash_date = c[0].date_input("Tarih", value=date.today())
            income = c[1].number_input("Gelir (TL)", min_value=0.0)
            expense = c[2].number_input("Gider (TL)", min_value=0.0)
            forecast = c[3].checkbox("Tahmini kayıt", value=True)
            if st.form_submit_button("Nakit akışı kaydet"):
                run_sql("""INSERT INTO cashflow(cash_date,income,expense,forecast)
                           VALUES(:d,:i,:e,:f)""",
                        {"d": cash_date, "i": income, "e": expense, "f": forecast})
                log_action("NAKİT AKIŞI", f"{cash_date}: +{income} / -{expense}")
                st.rerun()
        cash_df = load_df("cashflow", "cash_date")
        st.dataframe(cash_df, use_container_width=True, hide_index=True)
        
        if not cash_df.empty:
            cash_df["Net"] = pd.to_numeric(cash_df["income"], errors="coerce").fillna(0) - pd.to_numeric(cash_df["expense"], errors="coerce").fillna(0)
            cash_df["Kümülatif"] = cash_df["Net"].cumsum()
            
            fig_cash = go.Figure()
            fig_cash.add_trace(go.Bar(x=cash_df["cash_date"], y=cash_df["Net"], name="Aylık Net", marker_color='#0A192F'))
            fig_cash.add_trace(go.Scatter(x=cash_df["cash_date"], y=cash_df["Kümülatif"], mode='lines+markers', name="Kümülatif Bakiye", line=dict(color='#E63946', width=3)))
            fig_cash.update_layout(title="Nakit Akışı ve Kümülatif Bakiye", barmode='group')
            st.plotly_chart(fig_cash, use_container_width=True)

# ------------------------- Risk & payments -------------------------
elif menu == "⚠️ Risk & Ödeme":
    st.subheader("Risk Kayıtları ve Hakediş / Ödeme")
    tab_risk, tab_payment = st.tabs(["Riskler", "Hakediş ve Ödemeler"])
    with tab_risk:
        with st.form("risk_form", clear_on_submit=True):
            description = st.text_area("Risk açıklaması")
            c = st.columns(3)
            impact = c[0].selectbox("Etki", ["Düşük", "Orta", "Yüksek", "Kritik"])
            probability = c[1].selectbox("Olasılık", ["Düşük", "Orta", "Yüksek"])
            mitigation = c[2].text_area("Önlem / azaltma planı")
            if st.form_submit_button("Riski kaydet"):
                if description.strip():
                    run_sql("""INSERT INTO risk_register(description,impact,probability,mitigation)
                               VALUES(:d,:i,:p,:m)""",
                            {"d": description.strip(), "i": impact, "p": probability, "m": mitigation.strip()})
                    log_action("RİSK", description)
                    st.rerun()
                else:
                    st.error("Risk açıklaması zorunludur.")
        st.dataframe(df_risk, use_container_width=True, hide_index=True)
    with tab_payment:
        with st.form("payment_form", clear_on_submit=True):
            c = st.columns(4)
            wbs = c[0].selectbox("WBS", wbs_options() or ["GENEL"])
            amount = c[1].number_input("Tutar (TL)", min_value=0.0)
            payment_date = c[2].date_input("Tarih", value=date.today())
            status = c[3].selectbox("Durum", ["Hazırlanıyor", "Onayda", "Ödendi"])
            if st.form_submit_button("Hakediş / ödeme kaydet"):
                run_sql("""INSERT INTO progress_payment(wbs_code,amount,payment_date,status)
                           VALUES(:w,:a,:d,:s)""",
                        {"w": wbs, "a": amount, "d": payment_date, "s": status})
                log_action("HAKEDİŞ / ÖDEME", f"{wbs}: {amount} TL")
                st.rerun()
        st.dataframe(load_df("progress_payment", "payment_date DESC"), use_container_width=True, hide_index=True)

# ------------------------- Import / Reports -------------------------
elif menu == "⚙️ Veri Aktarımı & Raporlar":
    st.subheader("Excel Aktarımı ve Raporlar")
    with st.expander("Excel'den iş programı aktar"):
        st.caption("Excel sütun sırası: WBS | İş Kalemi | Süre | Bütçe. İlk satır başlık olmalıdır.")
        upload = st.file_uploader("Excel dosyası", type=["xlsx"])
        if upload:
            try:
                excel_df = pd.read_excel(upload)
                st.dataframe(excel_df.head(10), use_container_width=True)
                
                # Mükerrer WBS kontrolü
                if excel_df.iloc[:, 0].duplicated().any():
                    st.error("Excel dosyasında mükerrer WBS kodları var! Lütfen kontrol edin.")
                    st.write(excel_df[excel_df.iloc[:, 0].duplicated(keep=False)])
                elif st.button("Excel verilerini aktar"):
                    if excel_df.shape[1] < 4:
                        st.error("Excel dosyasında en az dört sütun bulunmalıdır.")
                    else:
                        clean_rows, errors = [], []
                        for idx, row in excel_df.iterrows():
                            w = str(row.iloc[0]).strip().upper()
                            name = str(row.iloc[1]).strip()
                            dur = pd.to_numeric(row.iloc[2], errors="coerce")
                            bud = pd.to_numeric(row.iloc[3], errors="coerce")
                            if w in {"NAN", "NONE"} or not valid_wbs(w) or not name or pd.isna(dur) or dur < 1 or pd.isna(bud) or bud < 0:
                                errors.append(f"{idx + 2}. satır: WBS, ad, süre veya bütçe hatalı.")
                            else:
                                clean_rows.append((w, name, int(dur), float(bud)))
                        if errors:
                            st.error("Aktarım durduruldu; hiçbir satır yazılmadı.")
                            st.write(errors[:30])
                        elif clean_rows:
                            with conn.session as session:
                                for w, name, dur, bud in clean_rows:
                                    session.execute(text("""
                                        INSERT INTO activities(wbs_code,task_name,duration,budget)
                                        VALUES(:w,:n,:d,:b)
                                        ON CONFLICT(wbs_code) DO UPDATE SET
                                        task_name=EXCLUDED.task_name,duration=EXCLUDED.duration,
                                        budget=EXCLUDED.budget,updated_at=CURRENT_TIMESTAMP
                                    """), {"w": w, "n": name, "d": dur, "b": bud})
                                session.commit()
                            log_action("EXCEL AKTARIMI", f"{len(clean_rows)} aktivite")
                            st.success(f"{len(clean_rows)} aktivite aktarıldı.")
                            st.rerun()
            except Exception as exc:
                st.error(f"Excel okunamadı: {exc}")

    if not df_act.empty:
        st.markdown("#### Excel raporu")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_act.to_excel(writer, sheet_name="Aktiviteler", index=False)
            df_dep.to_excel(writer, sheet_name="Bağımlılıklar", index=False)
            df_sat.to_excel(writer, sheet_name="Satın Alma", index=False)
            df_isc.to_excel(writer, sheet_name="İşçilik", index=False)
            df_overhead.to_excel(writer, sheet_name="Genel Giderler", index=False)
            load_df("progress_history").to_excel(writer, sheet_name="İlerleme Tarihçesi", index=False)
            load_df("risk_register").to_excel(writer, sheet_name="Riskler", index=False)
            docs_meta = run_sql("SELECT id, wbs_code, file_name, upload_date FROM documents ORDER BY upload_date DESC", fetch=True)
            docs_meta.to_excel(writer, sheet_name="Dokümanlar", index=False)
            
        st.download_button("📥 Tüm verileri Excel indir", output.getvalue(),
                           "EmirAsel_Proje_Raporu.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if not cpm_df.empty:
            evm = calculate_evm(cpm_df, df_sat, df_isc, df_overhead, project_start)
            if st.button("PDF yönetici raporunu hazırla"):
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=35, leftMargin=35, topMargin=35, bottomMargin=35)
                styles = getSampleStyleSheet()
                title_style = ParagraphStyle("EA_Title", parent=styles["Title"], textColor=colors.HexColor("#0A192F"), fontSize=16)
                story = [
                    Paragraph("EMİR ASEL İNŞAAT — YÖNETİCİ RAPORU", title_style),
                    Spacer(1, 8),
                    Paragraph(f"Rapor tarihi: {date.today().strftime('%d.%m.%Y')}", styles["Normal"]),
                    Paragraph(f"Proje başlangıcı: {project_start.strftime('%d.%m.%Y')}", styles["Normal"]),
                    Spacer(1, 14),
                ]
                report_rows = [
                    ["Gösterge", "Değer"],
                    ["Toplam bütçe (BAC)", f'{evm["BAC"]:,.2f} TL'],
                    ["Planlanan değer (PV)", f'{evm["PV"]:,.2f} TL'],
                    ["Kazanılmış değer (EV)", f'{evm["EV"]:,.2f} TL'],
                    ["Gerçekleşen maliyet (AC)", f'{evm["AC"]:,.2f} TL'],
                    ["CPI", f'{evm["CPI"]:.2f}'],
                    ["SPI", f'{evm["SPI"]:.2f}'],
                    ["Tahmini bitiş maliyeti (EAC)", f'{evm["EAC"]:,.2f} TL'],
                    ["Fiziksel ilerleme", f'%{cpm_df["Fiziksel İlerleme (%)"].mean():.1f}'],
                ]
                table = Table(report_rows, colWidths=[250, 220])
                table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0A192F")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), .5, colors.grey),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F1F4F7")]),
                    ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                    ("PADDING", (0, 0), (-1, -1), 7),
                ]))
                story.append(table)
                story.extend([Spacer(1, 16), Paragraph("Aktivite özeti", styles["Heading2"])])
                activity_rows = [["WBS", "İş kalemi", "Süre", "İlerleme", "Kritik"]]
                for _, r in cpm_df.iterrows():
                    activity_rows.append([str(r["WBS"]), Paragraph(str(r["İş Kalemi"]), styles["BodyText"]), str(r["Süre (Gün)"]), f'%{r["Fiziksel İlerleme (%)"]:.1f}', str(r["Kritik"])])
                activity_table = Table(activity_rows, repeatRows=1, colWidths=[45, 245, 50, 70, 55])
                activity_table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#183153")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), .4, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("PADDING", (0, 0), (-1, -1), 5),
                ]))
                story.append(activity_table)
                doc.build(story)
                
                pdf_data = buffer.getvalue()
                buffer.close()
                st.download_button("📄 PDF raporunu indir", pdf_data, "EmirAsel_Yonetici_Raporu.pdf", mime="application/pdf")

# ------------------------- Audit log -------------------------
elif menu == "🧾 İşlem Kayıtları":
    st.subheader("Kullanıcı İşlem Kayıtları")
    st.dataframe(load_df("audit_log", "log_date DESC"), use_container_width=True, hide_index=True)

st.sidebar.markdown("---")
st.sidebar.caption("Emir Asel İnşaat • Tek proje / tek kullanıcı")
