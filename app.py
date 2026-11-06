# -*- coding: utf-8 -*-
"""
EMİR ASEL İNŞAAT | PROJE KONTROL V8.5
- Excel V5 dosyasını okur.
- ÇŞB Birim Fiyat PDF'lerini (acikpoz) okur, Poz No eşleşenlere fiyat çeker.
- Logolu PDF ve Excel raporu üretir.
"""
from __future__ import annotations
from io import BytesIO
from datetime import date, datetime
import re, json, os, tempfile
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import cm
    PDF_OK = True
except ImportError:
    PDF_OK = False

try:
    from acikpoz import parse_catalog
    ACIKPOZ_OK = True
except ImportError:
    ACIKPOZ_OK = False

APP_VERSION = "8.5"

st.set_page_config(
    page_title="Emir Asel | Proje Kontrol V8.5",
    page_icon="🏗️", layout="wide", initial_sidebar_state="expanded",
)

# ============================================================
# TEMA
# ============================================================
NAVY = "#14263D"; NAVY_2 = "#0A192F"
RED = "#C83E4D"; RED_ACCENT = "#E63946"
TEAL = "#278C82"; GOLD = "#D6A84B"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap');
    html, body, [class*="css"] {{ font-family: 'Montserrat', sans-serif; }}
    .stApp {{ background: linear-gradient(180deg, #F8F9FA 0%, #FFFFFF 100%); }}
    .ea-header {{
        background: linear-gradient(135deg, {NAVY_2} 0%, {NAVY} 100%);
        padding: 26px 32px; border-radius: 16px; color: white;
        margin-bottom: 22px; box-shadow: 0 8px 24px rgba(10,25,47,0.22);
    }}
    .ea-header h1 {{ color: white !important; margin: 0; font-size: 26px; font-weight: 800; }}
    .ea-header .sub {{ color: #B8C5D6; font-size: 12px; margin-top: 6px; letter-spacing: 2px; }}
    section[data-testid="stSidebar"] {{ background: linear-gradient(180deg, {NAVY_2} 0%, {NAVY} 100%); }}
    section[data-testid="stSidebar"] * {{ color: white !important; }}
    div[data-testid="stMetric"] {{
        background: white; border-radius: 12px; padding: 16px 20px;
        box-shadow: 0 3px 10px rgba(10,25,47,0.06); border-left: 5px solid {RED};
    }}
    div[data-testid="stMetricValue"] {{ color: {NAVY} !important; font-weight: 800 !important; }}
    .stButton > button {{ background: linear-gradient(135deg, {RED_ACCENT} 0%, #A42A38 100%); color: white; border-radius: 10px; font-weight: 700; }}
    .section-title {{ color: {NAVY}; font-size: 18px; font-weight: 800; margin: 22px 0 10px 0; padding-left: 12px; border-left: 4px solid {RED_ACCENT}; }}
    .ea-footer {{ margin-top: 40px; padding: 20px 0; text-align: center; color: #94A3B8; font-size: 11px; border-top: 1px solid #E5EAF0; }}
</style>
""", unsafe_allow_html=True)

# ============================================================
# ŞEMALAR
# ============================================================
PLAN_COLS = ["WBS","Poz No","İş Kalemi","Ana Kategori","Alt Kategori","Birim","Miktar","Birim Fiyat","Planlanan Tutar","Başlangıç","Bitiş","Sorumlu","Açıklama"]
ACTUAL_COLS = ["WBS","Poz No","İş Kalemi","Ana Kategori","Birim","Gerçekleşen Miktar","Gerçekleşen Birim Fiyat","Gerçekleşen Tutar","Tarih","Tedarikçi","Belge No","Açıklama"]
SCHEDULE_COLS = ["WBS","Aktivite","Süre (gün)","Öncül WBS","Bağımlılık Tipi","Başlangıç","Bitiş","Fiili Başlangıç","Fiili Bitiş","Tamamlanma (%)","Ağırlık (%)","Sorumlu","Durum","Kanıt / Saha Notu"]
RISK_COLS = ["Risk ID","Risk Tanımı","Kategori","Olasılık (1-5)","Etki (1-5)","Risk Skoru","Önlem","Sorumlu","Durum","Tahmini Maliyet Etkisi"]
CASH_COLS = ["Dönem / Ay", "Planlanan Tahsilat", "Gerçekleşen Tahsilat", "Net Nakit Akışı", "Kümülatif Bakiye"]

COST_CATEGORIES = ["KABA İŞLER", "İNCE İŞLER / MEP", "DIŞ İŞLER", "GENEL GİDER / İŞÇİLİK",
                   "Mimarlık / Proje Gideri","Yapı Denetim Gideri","Yapı Analiz Laboratuvar Gideri","Hafriyat","Grobeton ve Beton",
                   "Toplam Kalıp + Demir İşçiliği","Demir / Donatı Malzemesi","Elektrik + Su İşçiliği","Elektrik Malzemeleri","Mekanik Tesisat",
                   "Duvar ve Sıva","Yalıtım","Şap ve Zemin","Seramik ve Kaplama","Kapı / Pencere / Doğrama","Cephe","Çatı","Boya",
                   "Peyzaj ve Çevre Düzenleme","Nakliye / Lojistik","Şantiye Genel Giderleri","İş Makinesi / Ekipman","İSG / Çevre","Diğer"]

# ============================================================
# YARDIMCI FONKSİYONLAR
# ============================================================
def safe_num(value, default=0.0):
    if pd.isna(value): return default
    if isinstance(value, (int, float, np.number)): return float(value)
    s = str(value).strip().replace("₺","").replace("$","").replace("€","").replace(" ","")
    if "," in s and "." in s:
        s = s.replace(".","").replace(",",".") if s.rfind(",")>s.rfind(".") else s.replace(",","")
    elif "," in s:
        s = s.replace(",",".")
    try: return float(s)
    except (TypeError, ValueError): return default

def norm_col(x):
    x = str(x).strip().lower()
    x = x.translate(str.maketrans("çğıöşüİ","cgiosui"))
    return re.sub(r"[^a-z0-9]+","",x)

def find_column(df, candidates):
    lookup = {norm_col(c): c for c in df.columns}
    for cand in candidates:
        if norm_col(cand) in lookup: return lookup[norm_col(cand)]
    for cand in candidates:
        key = norm_col(cand)
        for normalized, original in lookup.items():
            if key and (key in normalized or normalized in key): return original
    return None

def money(v, currency="₺"):
    try: return f"{currency} {float(v):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    except Exception: return f"{currency} 0,00"

def pct(v): return f"%{float(v):.1f}".replace(".",",")

def calc_plan_amount(df):
    df = df.copy()
    q = pd.to_numeric(df["Miktar"], errors="coerce").fillna(0)
    p = pd.to_numeric(df["Birim Fiyat"], errors="coerce").fillna(0)
    df["Planlanan Tutar"] = np.where(df["Planlanan Tutar"].isna() | (df["Planlanan Tutar"]==0), q*p, df["Planlanan Tutar"])
    return df

def calc_actual_amount(df):
    df = df.copy()
    q = pd.to_numeric(df["Gerçekleşen Miktar"], errors="coerce").fillna(0)
    p = pd.to_numeric(df["Gerçekleşen Birim Fiyat"], errors="coerce").fillna(0)
    df["Gerçekleşen Tutar"] = np.where(df["Gerçekleşen Tutar"].isna() | (df["Gerçekleşen Tutar"]==0), q*p, df["Gerçekleşen Tutar"])
    return df

def summarize_costs(plan, actual):
    p = plan.groupby("Ana Kategori", dropna=False)["Planlanan Tutar"].sum().rename("Bütçe")
    a = actual.groupby("Ana Kategori", dropna=False)["Gerçekleşen Tutar"].sum().rename("Gerçekleşen")
    result = pd.concat([p, a], axis=1).fillna(0).reset_index()
    result["Sapma"] = result["Gerçekleşen"] - result["Bütçe"]
    return result.sort_values("Bütçe", ascending=False)

def weighted_progress(schedule):
    if schedule.empty: return 0.0
    progress = pd.to_numeric(schedule["Tamamlanma (%)"], errors="coerce").fillna(0).clip(0,100)
    if "Ağırlık (%)" in schedule.columns:
        weights = pd.to_numeric(schedule["Ağırlık (%)"], errors="coerce").fillna(0).clip(lower=0)
        if weights.sum()>0: return float(np.average(progress, weights=weights))
    durations = pd.to_numeric(schedule["Süre (gün)"], errors="coerce").fillna(0).clip(lower=0)
    if durations.sum()<=0: return float(progress.mean()) if len(progress) else 0.0
    return float(np.average(progress, weights=durations))

# ============================================================
# ÇŞB PDF OKUMA (acikpoz)
# ============================================================
def extract_prices_from_pdf(pdf_bytes):
    if not ACIKPOZ_OK:
        return None, "acikpoz yüklü değil. requirements.txt'ye 'acikpoz' ekleyin."
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        result = parse_catalog(tmp_path)
        prices = {}
        for p in result.pozes:
            if getattr(p, "is_priced", False) and getattr(p, "fiyat", None) is not None:
                prices[str(p.poz_no).strip()] = float(p.fiyat)
        return prices, None
    except Exception as e:
        return None, str(e)
    finally:
        try: os.unlink(tmp_path)
        except: pass

# ============================================================
# EXCEL İÇE AKTARMA
# ============================================================
def read_uploaded_excel(upload):
    return pd.read_excel(upload, sheet_name=None, header=2)

def map_imported_sheet(df, kind):
    df = df.copy()
    if kind=="plan":
        mapping = {"WBS":["WBS","WBS Kodu","Poz No"], "Poz No":["Poz No","Poz"],
            "İş Kalemi":["İş Kalemi","Açıklama","İmalat"], "Ana Kategori":["Ana Kategori","Kategori"],
            "Birim":["Birim"], "Miktar":["Miktar","Metraj"], "Birim Fiyat":["Birim Fiyat","BF"],
            "Planlanan Tutar":["Planlanan Tutar","Toplam Tutar"]}
        out = pd.DataFrame()
        for t, c in mapping.items():
            col = find_column(df, c); out[t] = df[col] if col else np.nan
        out["Ana Kategori"] = out["Ana Kategori"].fillna("Diğer")
        out["Miktar"] = out["Miktar"].apply(safe_num); out["Birim Fiyat"] = out["Birim Fiyat"].apply(safe_num)
        out = calc_plan_amount(out)
        for c in PLAN_COLS:
            if c not in out.columns: out[c] = ""
        return out[PLAN_COLS]
    elif kind=="actual":
        mapping = {"WBS":["WBS Kodu","WBS"], "İş Kalemi":["İş Kalemi Adı","İş Kalemi"],
            "Gerçekleşen Miktar":["Fiili Miktar","Gerçekleşen Miktar"],
            "Gerçekleşen Birim Fiyat":["Fiili Birim Fiyat","Gerçekleşen Birim Fiyat"],
            "Tarih":["Tarih"], "Tedarikçi":["Tedarikçi / Firma","Tedarikçi"],
            "Belge No":["Fatura / Belge No","Belge No"]}
        out = pd.DataFrame()
        for t, c in mapping.items():
            col = find_column(df, c); out[t] = df[col] if col else np.nan
        out["Gerçekleşen Miktar"] = out["Gerçekleşen Miktar"].apply(safe_num)
        out["Gerçekleşen Birim Fiyat"] = out["Gerçekleşen Birim Fiyat"].apply(safe_num)
        out = calc_actual_amount(out)
        for c in ACTUAL_COLS:
            if c not in out.columns: out[c] = ""
        return out[ACTUAL_COLS]
    elif kind=="schedule":
        mapping = {"WBS":["WBS Kodu","WBS"], "Aktivite":["Aktivite Adı","Aktivite"],
            "Süre (gün)":["Süre (İş Günü)","Süre (gün)"], "Öncül WBS":["Öncül Faaliyet","Öncül WBS"],
            "Bağımlılık Tipi":["İlişki Tipi"], "Başlangıç":["Başlangıç Tarihi","Başlangıç"],
            "Bitiş":["Bitiş Tarihi","Bitiş"], "Tamamlanma (%)":["Gerçekleşen İlerleme %","Tamamlanma (%)"],
            "Sorumlu":["Sorumlu"], "Durum":["Durum"], "Kanıt / Saha Notu":["Saha Notu / Kanıt"]}
        out = pd.DataFrame()
        for t, c in mapping.items():
            col = find_column(df, c); out[t] = df[col] if col else np.nan
        out["Süre (gün)"] = out["Süre (gün)"].apply(safe_num)
        out["Tamamlanma (%)"] = out["Tamamlanma (%)"].apply(safe_num)
        out["Tamamlanma (%)"] = np.where(out["Tamamlanma (%)"] <= 1.0, out["Tamamlanma (%)"] * 100, out["Tamamlanma (%)"])
        for c in SCHEDULE_COLS:
            if c not in out.columns: out[c] = ""
        return out[SCHEDULE_COLS]
    elif kind=="risk":
        mapping = {"Risk ID":["Risk ID"], "Risk Tanımı":["Risk Tanımı"],
            "Kategori":["Kategori"], "Olasılık (1-5)":["Olasılık (1-5)"],
            "Etki (1-5)":["Etki (1-5)"], "Önlem":["Önleyici Faaliyet / Eylem Planı","Önlem"],
            "Sorumlu":["Sorumlu"], "Durum":["Durum"]}
        out = pd.DataFrame()
        for t, c in mapping.items():
            col = find_column(df, c); out[t] = df[col] if col else np.nan
        out["Olasılık (1-5)"] = out["Olasılık (1-5)"].apply(safe_num)
        out["Etki (1-5)"] = out["Etki (1-5)"].apply(safe_num)
        out["Risk Skoru"] = out["Olasılık (1-5)"] * out["Etki (1-5)"]
        for c in RISK_COLS:
            if c not in out.columns: out[c] = ""
        return out[RISK_COLS]
    elif kind=="cash":
        mapping = {"Dönem / Ay":["Dönem / Ay","Ay"], "Planlanan Tahsilat":["Planlanan Tahsilat (TL)"],
            "Gerçekleşen Tahsilat":["Gerçekleşen Tahsilat (TL)"],
            "Net Nakit Akışı":["Net Nakit Akışı (TL)"], "Kümülatif Bakiye":["Kümülatif Bakiye (TL)"]}
        out = pd.DataFrame()
        for t, c in mapping.items():
            col = find_column(df, c); out[t] = df[col] if col else np.nan
        for c in CASH_COLS:
            if c not in out.columns: out[c] = 0
        return out[CASH_COLS]
    return df

# ============================================================
# DEMO VERİ
# ============================================================
def demo_data():
    plan = pd.DataFrame([
        ["1.1","01.001","Mimari Projeler","Mimarlık / Proje Gideri","Proje","LS",1,850000,850000,"","","",""],
        ["2.1","04.001","Temel Betonarme","Grobeton ve Beton","Beton","m³",450,5200,2340000,"","","",""],
        ["3.1","05.001","Elektrik Tesisat","Elektrik Malzemeleri","MEP","LS",1,310000,310000,"","","",""],
    ], columns=PLAN_COLS)
    actual = pd.DataFrame([
        ["1.1","01.001","Mimari Projeler","Mimarlık / Proje Gideri","LS",1,800000,800000,"2026-01-15","Proje Ofisi","F-001",""],
    ], columns=ACTUAL_COLS)
    schedule = pd.DataFrame([
        ["1.1","Proje ve Ruhsat",40,"","FS","2026-01-05","2026-02-15","","",100,40,"Proje","Tamamlandı",""],
        ["2.1","Temel İmalatları",45,"1.1","FS","2026-02-16","2026-04-01","","",65,45,"Kaba Yapı","Devam ediyor",""],
    ], columns=SCHEDULE_COLS)
    risk = pd.DataFrame([
        ["R-01","Beton fiyat artışı","Finansal",4,5,20,"Alternatif tedarikçi","Satın Alma","Açık",0],
    ], columns=RISK_COLS)
    cash = pd.DataFrame([
        ["2026-01", 350000, 350000, 0, 350000],
        ["2026-02", 750000, 720000, -30000, 1070000],
        ["2026-03", 950000, 980000, 30000, 2050000],
    ], columns=CASH_COLS)
    return plan, actual, schedule, risk, cash

# ============================================================
# OTURUM
# ============================================================
if "plan" not in st.session_state:
    p, a, s, r, c = demo_data()
    st.session_state.plan, st.session_state.actual, st.session_state.schedule, st.session_state.risk, st.session_state.cash = p, a, s, r, c

if "project_name" not in st.session_state:
    st.session_state.project_name = "Emir Asel Konut Projesi"

# ============================================================
# HEADER
# ============================================================
st.markdown(f"""
<div class="ea-header">
  <h1>🏗️ EMİR ASEL İNŞAAT</h1>
  <div class="sub">PROJE KONTROL VE MALİYET YÖNETİM PANELİ · V{APP_VERSION} · {datetime.now():%d.%m.%Y %H:%M}</div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    try:
        st.image("logo.png", use_column_width=True)
    except Exception:
        st.markdown("### 🏗️ EMİR ASEL")
        
    MENU = ["Yönetici Dashboard","Proje Künyesi (Excel Yükle)","Bütçe ve Metraj","Gerçekleşen Maliyet",
            "İş Programı / CPM","EVM Performans","Risk Yönetimi","Nakit Akışı (S-Curve)","Raporlar"]
    page = st.radio("MODÜLLER", MENU)
    st.divider()
    project_name = st.text_input("Proje Adı", st.session_state.project_name)
    company = st.text_input("Firma", "Emir Asel İnşaat")

# ============================================================
# VERİ HAZIRLA
# ============================================================
plan = st.session_state.plan.copy()
actual = st.session_state.actual.copy()
schedule = st.session_state.schedule.copy()
risk = st.session_state.risk.copy()
cash = st.session_state.cash.copy()

plan = calc_plan_amount(plan)
actual = calc_actual_amount(actual)
budget = float(plan["Planlanan Tutar"].sum())
actual_cost = float(actual["Gerçekleşen Tutar"].sum())
cost_summary = summarize_costs(plan, actual)
progress = weighted_progress(schedule)

ev = budget * (progress / 100) if budget > 0 else 0
pv = budget * 0.5
cpi = ev / actual_cost if actual_cost > 0 else 0
spi = ev / pv if pv > 0 else 0
eac = budget / cpi if cpi > 0 else 0

# ============================================================
# DASHBOARD
# ============================================================
if page == "Yönetici Dashboard":
    st.title("Yönetici Dashboard")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Toplam Bütçe (BAC)", money(budget))
    c2.metric("Gerçekleşen (AC)", money(actual_cost), delta=money(actual_cost-budget) if budget else None, delta_color="inverse")
    c3.metric("Kazanılan Değer (EV)", money(ev))
    c4.metric("Tahmini Bitiş (EAC)", money(eac))
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("CPI", f"{cpi:.2f}"); m2.metric("SPI", f"{spi:.2f}")
    m3.metric("İlerleme", pct(progress)); m4.metric("Kritik Risk", len(risk[risk["Risk Skoru"] >= 15]) if not risk.empty else 0)
    st.markdown('<div class="section-title">📊 Kategori Bazında Bütçe / Gerçekleşen</div>', unsafe_allow_html=True)
    if not cost_summary.empty:
        fig = px.bar(cost_summary, x="Ana Kategori", y=["Bütçe","Gerçekleşen"], barmode="group", color_discrete_sequence=[NAVY, RED])
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# EXCEL YÜKLEME
# ============================================================
elif page == "Proje Künyesi (Excel Yükle)":
    st.title("Proje Künyesi ve Excel Entegrasyonu")
    st.info("V5 Excel dosyanızı yükleyin. Sistem otomatik olarak Bütçe, Gerçekleşen, İş Programı, Risk ve Nakit Akışı sayfalarını okuyacaktır.")
    uploaded = st.file_uploader("Emir_Asel_Proje_Kontrol_V5.xlsx yükleyin", type=["xlsx"])
    if uploaded:
        try:
            raw_sheets = read_uploaded_excel(uploaded)
            st.success(f"Bulunan Sayfalar: {list(raw_sheets.keys())}")
            plan_df, actual_df, sched_df, risk_df, cash_df = None, None, None, None, None
            for name, df in raw_sheets.items():
                cols = [str(c).lower() for c in df.columns]
                if "miktar" in cols and "birim fiyat" in cols: plan_df = df
                if "fiili miktar" in cols or "gerçekleşen miktar" in cols: actual_df = df
                if "aktivite adı" in cols or "süre (iş günü)" in cols: sched_df = df
                if "risk id" in cols: risk_df = df
                if "planlanan tahsilat" in cols or "dönem / ay" in cols: cash_df = df
            if st.button("Verileri Sisteme Aktar", type="primary"):
                if plan_df is not None: st.session_state.plan = map_imported_sheet(plan_df, "plan")
                if actual_df is not None: st.session_state.actual = map_imported_sheet(actual_df, "actual")
                if sched_df is not None: st.session_state.schedule = map_imported_sheet(sched_df, "schedule")
                if risk_df is not None: st.session_state.risk = map_imported_sheet(risk_df, "risk")
                if cash_df is not None: st.session_state.cash = map_imported_sheet(cash_df, "cash")
                st.success("Aktarıldı!"); st.rerun()
        except Exception as e:
            st.error(f"Excel okunurken hata: {e}")

# ============================================================
# BÜTÇE (ÇŞB PDF Entegrasyonu)
# ============================================================
elif page == "Bütçe ve Metraj":
    st.title("Bütçe ve Metraj")
    st.markdown('<div class="section-title">📄 ÇŞB Birim Fiyat PDF\'inden Fiyat Çek</div>', unsafe_allow_html=True)
    if not ACIKPOZ_OK:
        st.warning("`acikpoz` kütüphanesi yüklü değil. `requirements.txt` dosyanıza `acikpoz` ekleyin.")
    else:
        st.info("ÇŞB sitesinden indirdiğiniz güncel Birim Fiyat PDF'ini yükleyin. Poz No eşleşen satırlara otomatik fiyat çeker.")
        pdf_file = st.file_uploader("Birim Fiyat PDF'i (ör: bf2026.pdf)", type=["pdf"], key="csb_pdf")
        if pdf_file is not None:
            if st.button("PDF'i Oku ve Fiyatları Çek", type="primary"):
                with st.spinner("PDF okunuyor..."):
                    prices, err = extract_prices_from_pdf(pdf_file.getvalue())
                if err: st.error(f"Hata: {err}")
                elif prices:
                    st.success(f"PDF'ten {len(prices)} adet fiyat okundu!")
                    updated = 0
                    for idx, row in plan.iterrows():
                        poz = str(row.get("Poz No", "")).strip()
                        if poz in prices:
                            plan.at[idx, "Birim Fiyat"] = prices[poz]
                            updated += 1
                    if updated > 0:
                        st.session_state.plan = calc_plan_amount(plan)
                        st.success(f"✅ {updated} satır güncellendi!"); st.rerun()
                    else:
                        st.warning("PDF'teki pozlar tablodaki Poz No ile eşleşmedi.")
                else:
                    st.warning("PDF'te fiyatlı poz bulunamadı.")
    st.divider()
    edited = st.data_editor(plan, num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={"Ana Kategori": st.column_config.SelectboxColumn(options=COST_CATEGORIES)})
    if st.button("Tabloyu Kaydet"):
        st.session_state.plan = calc_plan_amount(edited); st.success("Kaydedildi."); st.rerun()

# ============================================================
# GERÇEKLEŞEN
# ============================================================
elif page == "Gerçekleşen Maliyet":
    st.title("Gerçekleşen Maliyet")
    edited = st.data_editor(actual, num_rows="dynamic", use_container_width=True, hide_index=True)
    if st.button("Kaydet"):
        st.session_state.actual = calc_actual_amount(edited); st.success("Kaydedildi."); st.rerun()

# ============================================================
# İŞ PROGRAMI
# ============================================================
elif page == "İş Programı / CPM":
    st.title("İş Programı / CPM")
    edited = st.data_editor(schedule, num_rows="dynamic", use_container_width=True, hide_index=True)
    if st.button("Kaydet"):
        st.session_state.schedule = edited; st.success("Kaydedildi."); st.rerun()
    if not schedule.empty:
        st.markdown('<div class="section-title">📅 Gantt Şeması</div>', unsafe_allow_html=True)
        df_gantt = schedule.copy()
        df_gantt["Başlangıç"] = pd.to_datetime(df_gantt["Başlangıç"], errors="coerce")
        df_gantt["Bitiş"] = pd.to_datetime(df_gantt["Bitiş"], errors="coerce")
        df_gantt = df_gantt.dropna(subset=["Başlangıç", "Bitiş"])
        if not df_gantt.empty:
            fig = px.timeline(df_gantt, x_start="Başlangıç", x_end="Bitiş", y="Aktivite", color="Tamamlanma (%)", color_continuous_scale="Blues")
            fig.update_yaxes(autorange="reversed")
            st.plotly_chart(fig, use_container_width=True)

# ============================================================
# EVM
# ============================================================
elif page == "EVM Performans":
    st.title("EVM Performans")
    st.metric("BAC", money(budget)); st.metric("EV", money(ev))
    st.metric("AC", money(actual_cost)); st.metric("CPI", f"{cpi:.2f}"); st.metric("SPI", f"{spi:.2f}")
    fig = go.Figure(data=[
        go.Bar(name='BAC', x=['Bütçe'], y=[budget], marker_color=NAVY),
        go.Bar(name='EV', x=['Kazanılan'], y=[ev], marker_color=TEAL),
        go.Bar(name='AC', x=['Gerçekleşen'], y=[actual_cost], marker_color=RED)
    ])
    fig.update_layout(title="EVM Karşılaştırma", barmode='group')
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# RİSK
# ============================================================
elif page == "Risk Yönetimi":
    st.title("Risk Yönetimi")
    edited = st.data_editor(risk, num_rows="dynamic", use_container_width=True, hide_index=True)
    if st.button("Kaydet"):
        st.session_state.risk = edited; st.success("Kaydedildi."); st.rerun()
    if not risk.empty:
        st.markdown('<div class="section-title">⚠️ Risk Matrisi</div>', unsafe_allow_html=True)
        fig = px.scatter(risk, x="Olasılık (1-5)", y="Etki (1-5)", size="Risk Skoru", color="Risk Skoru",
                         hover_name="Risk Tanımı", text="Risk ID", color_continuous_scale="RdYlGn_r",
                         range_x=[0.5,5.5], range_y=[0.5,5.5])
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# NAKİT AKIŞI
# ============================================================
elif page == "Nakit Akışı (S-Curve)":
    st.title("Nakit Akışı (S-Curve)")
    edited = st.data_editor(cash, num_rows="dynamic", use_container_width=True, hide_index=True)
    if st.button("Kaydet"):
        st.session_state.cash = edited; st.success("Kaydedildi."); st.rerun()
    if not cash.empty:
        st.markdown('<div class="section-title">📈 S-Curve Grafiği</div>', unsafe_allow_html=True)
        cash["Kümülatif Bakiye"] = pd.to_numeric(cash["Gerçekleşen Tahsilat"], errors="coerce").fillna(0).cumsum()
        fig = go.Figure()
        fig.add_trace(go.Bar(x=cash["Dönem / Ay"], y=cash["Planlanan Tahsilat"], name="Planlanan", marker_color=NAVY))
        fig.add_trace(go.Bar(x=cash["Dönem / Ay"], y=cash["Gerçekleşen Tahsilat"], name="Gerçekleşen", marker_color=TEAL))
        fig.add_trace(go.Scatter(x=cash["Dönem / Ay"], y=cash["Kümülatif Bakiye"], name="Kümülatif", yaxis="y2", line=dict(color=RED, width=4)))
        fig.update_layout(title="Nakit Akışı", xaxis=dict(title="Dönem"),
                          yaxis2=dict(title="Kümülatif", overlaying="y", side="right"),
                          barmode="group")
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# RAPORLAR
# ============================================================
elif page == "Raporlar":
    st.title("Raporlar ve Dışa Aktarım")
    if EXCEL_OK:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            plan.to_excel(writer, sheet_name="Bütçe", index=False)
            actual.to_excel(writer, sheet_name="Gerçekleşen", index=False)
            schedule.to_excel(writer, sheet_name="İş Programı", index=False)
            risk.to_excel(writer, sheet_name="Riskler", index=False)
            cash.to_excel(writer, sheet_name="Nakit Akışı", index=False)
        st.download_button("📥 Excel Raporu İndir", data=output.getvalue(),
            file_name=f"EmirAsel_{datetime.now():%Y%m%d_%H%M}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if PDF_OK and os.path.exists("logo.png"):
        if st.button("📄 Logolu PDF Raporu Oluştur"):
            buf = BytesIO()
            doc = SimpleDocTemplate(buf, pagesize=A4)
            story = [RLImage("logo.png", width=5*cm, height=3*cm),
                     Paragraph(f"<b>EMİR ASEL İNŞAAT</b>", getSampleStyleSheet()['Title']),
                     Paragraph(f"Proje: {project_name}", getSampleStyleSheet()['Normal']),
                     Paragraph(f"Toplam Bütçe: {money(budget)}", getSampleStyleSheet()['Normal']),
                     Paragraph(f"Gerçekleşen: {money(actual_cost)}", getSampleStyleSheet()['Normal']),
                     Paragraph(f"İlerleme: {pct(progress)}", getSampleStyleSheet()['Normal'])]
            doc.build(story); buf.seek(0)
            st.download_button("📥 PDF İndir", data=buf.getvalue(),
                file_name=f"EmirAsel_{datetime.now():%Y%m%d_%H%M}.pdf", mime="application/pdf")

st.markdown(f'<div class="ea-footer"><strong>EMİR ASEL İNŞAAT</strong> · V{APP_VERSION}</div>', unsafe_allow_html=True)
