import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import datetime, timedelta
import io
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from sqlalchemy import text
from PIL import Image
from streamlit_option_menu import option_menu

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

plt.rcParams['font.sans-serif'] = ['DejaVu Sans']

st.set_page_config(page_title="Emir Asel İnşaat — Proje Yönetimi", page_icon="🏗️", layout="wide")

# ============================================================
# PROFESYONEL TASARIM (CSS) - Logoya Uygun Kırmızı/Lacivert Tema
# ============================================================
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
    div[data-testid="stInfo"] { background-color: #E3F2FD; border-left: 5px solid #1976D2; border-radius: 5px; }
    section[data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #EAEAEA; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# LOGO VE VERİTABANI BAĞLANTISI
# ============================================================
logo = None
for dosya in ["logo.png", "logo.png.webp", "logo.webp", "logo.jpg", "logo.jpeg"]:
    if os.path.exists(dosya):
        try:
            logo = Image.open(dosya); break
        except Exception: pass
if logo: st.sidebar.image(logo, use_container_width=True)

try: conn = st.connection("neon", type="sql")
except Exception as e:
    st.error(f"Veritabanı bağlantısı kurulamadı: {e}"); st.stop()

def init_db():
    try:
        with conn.session as s:
            s.execute(text("CREATE TABLE IF NOT EXISTS ilerleme (wbs_kodu TEXT PRIMARY KEY, ilerleme_yuzdesi INTEGER DEFAULT 0)"))
            s.execute(text("CREATE TABLE IF NOT EXISTS satin_alma (id SERIAL PRIMARY KEY, malzeme TEXT, tedarikci TEXT, tutar NUMERIC, durum TEXT, tarih TEXT)"))
            s.execute(text("CREATE TABLE IF NOT EXISTS iscilik (id SERIAL PRIMARY KEY, ekip TEXT, is_tanimi TEXT, adet INTEGER, yevmiye NUMERIC, tarih TEXT)"))
            s.commit()
    except Exception as e: st.warning(f"Tablo oluşturulamadı: {e}")

def ilerleme_kaydet(wbs, yuzde):
    with conn.session as s:
        s.execute(text("INSERT INTO ilerleme (wbs_kodu, ilerleme_yuzdesi) VALUES (:wbs, :yuzde) ON CONFLICT(wbs_kodu) DO UPDATE SET ilerleme_yuzdesi = :yuzde"), {"wbs": wbs, "yuzde": yuzde})
        s.commit()

def ilerleme_yukle():
    try:
        df = conn.query("SELECT wbs_kodu, ilerleme_yuzdesi FROM ilerleme", ttl=0)
        return dict(zip(df['wbs_kodu'], df['ilerleme_yuzdesi']))
    except Exception: return {}

init_db()

# ============================================================
# VARSAYILAN VERİ & CPM MOTORU
# ============================================================
@st.cache_data
def varsayilan_veri():
    veriler = [
        ["A01","Hafriyat kazı",15,"","FS", 95000], ["A02","Grobeton C15/20",5,"A01","FS", 61000], ["A03","Blokaj taş dolgu",3,"A01","FS", 7200],
        ["A04","Sıkıştırılmış dolgu",2,"A03","FS", 19000], ["A05","Zemin etüdü",7,"","FS", 28000], ["A06","Resmi işlemler (Ruhsat)",30,"","FS", 65000],
        ["B01","Radye temel betonu C30/37",20,"A02,A04","FS", 247000], ["B02","Radye altı membran yalıtım",5,"A02","FS", 33500], ["B03","Radye donatı Ø8",5,"B01","FS", 13600],
        ["B04","Radye donatı Ø10",5,"B01","FS", 21400], ["B05","Radye donatı Ø12",5,"B01","FS", 31200], ["B06","Radye donatı Ø14",5,"B01","FS", 42900],
        ["B07","Radye donatı Ø16",5,"B01","FS", 23400], ["B08","Radye donatı Ø20",5,"B01","FS", 7800], ["B09","Radye temel kalıbı",10,"B03,B04,B05,B06,B07,B08","FS", 20900],
        ["B10","Demir bağlama teli + sarf",2,"B01","FS", 9000], ["C01","Subasman perde betonu C30",10,"B09","FS", 43200], ["C02","Subasman donatı Ø12",3,"C01","FS", 35100],
        ["C03","Subasman donatı Ø14",3,"C01","FS", 23400], ["C04","Subasman donatı Ø16",3,"C01","FS", 9100], ["C05","Subasman kalıbı",5,"C02,C03,C04","FS", 39000],
        ["C06","Subasman dış yalıtım (XPS)",4,"C05","FS", 23400], ["C07","Drenaj borusu + çakıl",3,"C06","FS", 10000], ["D01","Zemin kat kolon betonu C30",15,"C07","FS", 28800],
        ["D02","Zemin kat perde betonu C30",10,"C07","FS", 21600], ["D03","Kolon donatı Ø16",4,"D01","FS", 35100], ["D04","Kolon donatı Ø20",4,"D01","FS", 27300],
        ["D05","Perde donatı Ø12",3,"D02","FS", 11700], ["D06","Perde donatı Ø14",3,"D02","FS", 7900], ["D07","Kolon kalıbı",8,"D03,D04","FS", 21400], ["D08","Perde kalıbı",6,"D05,D06","FS", 15600],
    ]
    df = pd.DataFrame(veriler, columns=["WBS","İş Kalemi","Süre (Gün)","Önceki Aktivite","Bağımlılık Tipi","Bütçe (TL)"])
    for col in ["Sorumlu","Plan Başlangıç","Plan Bitiş","Fiili Başlangıç","Fiili Bitiş", "Plan İlerleme %","Fiili İlerleme %","Ağırlık %","Ağırlıklı İlerleme %", "Durum","Gecikme (gün)","Kanıt / Saha Notu"]: df[col] = ""
    return df[["WBS","İş Kalemi","Sorumlu","Plan Başlangıç","Plan Bitiş", "Fiili Başlangıç","Fiili Bitiş","Plan İlerleme %","Fiili İlerleme %", "Ağırlık %","Ağırlıklı İlerleme %","Durum","Gecikme (gün)", "Kanıt / Saha Notu","Süre (Gün)","Önceki Aktivite","Bağımlılık Tipi","Bütçe (TL)"]]

@st.cache_data
def cpm_hesapla(df_raw, baslangic="2026-01-15"):
    akt = {}
    for _, row in df_raw.iterrows():
        row_dict = {str(k).strip().lower(): v for k, v in row.items()}
        wbs = str(row_dict.get('wbs kodu', row_dict.get('wbs', ''))).strip()
        if not wbs or wbs.lower() in ('nan','none',''): continue
        task = str(row_dict.get('i̇ş kalemi', row_dict.get('iş kalemi', 'İsimsiz')))
        try: dur = int(float(row_dict.get('süre (gün)', row_dict.get('süre', 1)))); dur = max(1, dur)
        except: dur = 1
        preds = [p.strip() for p in str(row_dict.get('önceki aktivite', row_dict.get('önceki', ''))).split(',') if p.strip() and p.strip().lower() != 'nan']
        dtype = str(row_dict.get('bağımlılık tipi', row_dict.get('tip', 'FS'))).upper()
        if dtype not in ["FS","SS","FF","SF"]: dtype = "FS"
        try: butce = float(row_dict.get('bütçe (tl)', row_dict.get('bütçe tutarı', row_dict.get('bütçe', 0))))
        except: butce = 0.0
        akt[wbs] = {'WBS':wbs,'Task':task,'Dur':dur, 'Preds':preds,'Type':dtype,'Succ':[], 'Butce': butce}
    wbs_set = set(akt.keys())
    for w, a in akt.items():
        valid = [p for p in a['Preds'] if p in wbs_set]
        for p in valid: akt[p]['Succ'].append({'wbs':w,'type':a['Type']})
        a['Preds'] = valid
    in_d = {w: len(a['Preds']) for w,a in akt.items()}
    q = [w for w,d in in_d.items() if d==0]; topo = []
    while q:
        c = q.pop(0); topo.append(c)
        for s in akt[c]['Succ']:
            in_d[s['wbs']] -= 1
            if in_d[s['wbs']]==0: q.append(s['wbs'])
    es, ef = {}, {}
    for w in topo:
        a = akt[w]; dur = a['Dur']
        e_start = 1 if not a['Preds'] else max([ef[p]+1 if a['Type']=="FS" else es[p] if a['Type']=="SS" else ef[p] if a['Type']=="FF" else es[p] for p in a['Preds']])
        es[w] = max(1, e_start); ef[w] = es[w]+dur-1
    proje_bitis = max(ef.values()) if ef else 1
    ls, lf = {}, {}
    for w in reversed(topo):
        a = akt[w]; dur = a['Dur']
        if not a['Succ']: lf[w] = proje_bitis
        else: lf[w] = min([ls[s['wbs']]-1 if s['type']=="FS" else ls[s['wbs']] if s['type']=="SS" else lf[s['wbs']] if s['type']=="FF" else lf[s['wbs']] for s in a['Succ']])
        ls[w] = lf[w]-dur+1
    base = datetime.strptime(baslangic, "%Y-%m-%d"); toplam = sum(a['Dur'] for a in akt.values()); toplam_butce = sum(a['Butce'] for a in akt.values())
    rows = []
    for w in topo:
        a = akt[w]; dur = a['Dur']; slack = ls[w]-es[w]
        rows.append({"WBS Kodu": w, "İş Kalemi": a['Task'], "Süre (Gün)": dur, "Önceki": ",".join(a['Preds']), "Bağlantı": a['Type'], "ES": es[w], "EF": ef[w], "LS": ls[w], "LF": lf[w], "Slack": slack, "Kritik Mi?": "EVET" if slack==0 else "HAYIR", "Plan Başlangıç": (base+timedelta(days=es[w]-1)).strftime("%d.%m.%Y"), "Plan Bitiş": (base+timedelta(days=ef[w]-1)).strftime("%d.%m.%Y"), "Ağırlık (%)": round((dur/toplam)*100,2) if toplam else 0, "Bütçe (TL)": a['Butce']})
    return pd.DataFrame(rows), proje_bitis, toplam_butce

# ============================================================
# ANA UYGULAMA
# ============================================================
st.markdown('<div class="header-container"><h1 class="header-title">🏗️ EMİR ASEL İNŞAAT — PROJE YÖNETİM PANELİ</h1></div>', unsafe_allow_html=True)

# Yan Menü Yükleme
st.sidebar.header("📂 Proje Verisi Yükle")
uploaded_excel = st.sidebar.file_uploader("Excel Yükle (WBS, Süre, Bütçe)", type=["xlsx"])
df_veri = varsayilan_veri()

if uploaded_excel is not None:
    try:
        df_yuklenen = pd.read_excel(uploaded_excel, sheet_name="05_IS_PROGRAMI", header=3)
        df_yuklenen.columns = [str(c).strip() for c in df_yuklenen.columns]
        df_veri = pd.DataFrame()
        df_veri['WBS'] = df_yuklenen.iloc[:, 0]; df_veri['İş Kalemi'] = df_yuklenen.iloc[:, 1]
        for col in ['Süre (Gün)', 'Süre']:
            if col in df_yuklenen.columns: df_veri['Süre (Gün)'] = df_yuklenen[col]; break
        else: df_veri['Süre (Gün)'] = 1
        for col in ['Önceki Aktivite', 'Önceki']:
            if col in df_yuklenen.columns: df_veri['Önceki Aktivite'] = df_yuklenen[col]; break
        else: df_veri['Önceki Aktivite'] = ""
        for col in ['Bağımlılık Tipi', 'Tip']:
            if col in df_yuklenen.columns: df_veri['Bağımlılık Tipi'] = df_yuklenen[col]; break
        else: df_veri['Bağımlılık Tipi'] = "FS"
        for col in ['Bütçe (TL)', 'Bütçe Tutarı', 'Bütçe']:
            if col in df_yuklenen.columns: df_veri['Bütçe (TL)'] = df_yuklenen[col]; break
        else: df_veri['Bütçe (TL)'] = 0.0
        for col in ["Sorumlu","Plan Başlangıç","Plan Bitiş","Fiili Başlangıç","Fiili Bitiş", "Plan İlerleme %","Fiili İlerleme %","Ağırlık %","Ağırlıklı İlerleme %", "Durum","Gecikme (gün)","Kanıt / Saha Notu"]: df_veri[col] = ""
        df_veri = df_veri[["WBS","İş Kalemi","Sorumlu","Plan Başlangıç","Plan Bitiş", "Fiili Başlangıç","Fiili Bitiş","Plan İlerleme %","Fiili İlerleme %", "Ağırlık %","Ağırlıklı İlerleme %","Durum","Gecikme (gün)", "Kanıt / Saha Notu","Süre (Gün)","Önceki Aktivite","Bağımlılık Tipi","Bütçe (TL)"]]
        st.sidebar.success("✅ Excel yüklendi!")
    except Exception as e: st.sidebar.error(f"❌ Hata: {e}")

df_cpm, proje_suresi, toplam_butce = cpm_hesapla(df_veri)
kayitli_ilerleme = ilerleme_yukle()
evm_baslangic = df_cpm[df_cpm["Bütçe (TL)"] > 0][["WBS Kodu", "İş Kalemi", "Bütçe (TL)"]].copy()
evm_baslangic["İlerleme %"] = evm_baslangic["WBS Kodu"].map(kayitli_ilerleme).fillna(0).astype(int)

with st.sidebar:
    st.header("📊 Proje Özeti")
    st.metric("Toplam Süre", f"{proje_suresi} gün")
    st.metric("Kritik Akt.", len(df_cpm[df_cpm["Kritik Mi?"]=="EVET"]))
    st.metric("Toplam Bütçe", f"{toplam_butce:,.0f} TL")

# Sekmeler (Profesyonel Görünüm)
secili_sekme = option_menu(
    menu_title=None, options=["Dashboard", "CPM & Gantt", "EVM & Maliyet", "Satın Alma", "İşçilik"],
    icons=["speedometer2", "diagram-3", "cash-coin", "cart", "people"], menu_icon="cast", default_index=0, orientation="horizontal",
    styles={"container": {"padding": "0!important", "background-color": "#FFFFFF", "border-radius": "10px", "box-shadow": "0 2px 4px rgba(0,0,0,0.05)"}, "icon": {"color": "#1976D2", "font-size": "16px"}, "nav-link": {"font-size": "15px", "text-align": "center", "margin": "0px", "--hover-color": "#F0F2F6", "font-weight": "500"}, "nav-link-selected": {"background-color": "#E63946", "color": "white"}}
)

if secili_sekme == "Dashboard":
    guncel_ilerleme = ilerleme_yukle(); toplam_pv, toplam_ev, toplam_ac = 0, 0, 0; en_pahali_kalemler = []
    for _, row in df_cpm.iterrows():
        butce = row["Bütçe (TL)"]
        if butce > 0:
            ilerleme = guncel_ilerleme.get(row["WBS Kodu"], 0) / 100.0; toplam_pv += butce; ev = butce * ilerleme; toplam_ev += ev; toplam_ac += ev * 1.10
            en_pahali_kalemler.append({"İş Kalemi": row["İş Kalemi"], "Bütçe": butce, "İlerleme": ilerleme * 100})
    c1, c2, c3, c4 = st.columns(4)
    genel_ilerleme = (toplam_ev / toplam_pv * 100) if toplam_pv > 0 else 0
    c1.metric("Genel İlerleme", f"%{genel_ilerleme:.1f}"); c2.metric("Toplam Bütçe", f"{toplam_pv:,.0f} TL"); c3.metric("Hakediş (EV)", f"{toplam_ev:,.0f} TL")
    cpi = toplam_ev / toplam_ac if toplam_ac > 0 else 1.0; c4.metric("CPI", f"{cpi:.2f}", delta=f"{cpi-1:.2f}", delta_color="inverse" if cpi < 1 else "normal")
    st.markdown("---"); g1, g2 = st.columns(2)
    with g1:
        st.markdown("### 🏗️ Disiplinlere Göre Bütçe")
        df_bg = df_cpm[df_cpm["Bütçe (TL)"] > 0].copy(); df_bg['Grup'] = df_bg['WBS Kodu'].str[0]
        bo = df_bg.groupby(df_bg['Grup'].map({'A': 'Hazırlık', 'B': 'Temel', 'C': 'Subasman', 'D': 'Üst Yapı'}).fillna('Diğer'))['Bütçe (TL)'].sum().reset_index()
        fig1, ax1 = plt.subplots(figsize=(7, 5)); ax1.pie(bo['Bütçe (TL)'], labels=bo['Grup'], autopct='%1.1f%%', startangle=90, colors=plt.cm.Set3.colors); ax1.axis('equal'); st.pyplot(fig1); plt.close(fig1)
    with g2:
        st.markdown("### 💰 En Yüksek 5 Kalem")
        en_pahali_kalemler.sort(key=lambda x: x["Bütçe"], reverse=True); top5 = en_pahali_kalemler[:5]
        fig2, ax2 = plt.subplots(figsize=(7, 5)); bars = ax2.barh([k["İş Kalemi"][:20]+"..." for k in top5], [k["Bütçe"] for k in top5], color='#1976D2', alpha=0.7, edgecolor='black')
        for bar, kalem in zip(bars, top5): ax2.text(bar.get_width() / 2, bar.get_y() + bar.get_height()/2, f'%{kalem["İlerleme"]:.0f}', ha='center', va='center', color='white', fontweight='bold', fontsize=9)
        ax2.set_xlabel("Bütçe (TL)"); ax2.invert_yaxis(); st.pyplot(fig2); plt.close(fig2)
    st.markdown("---"); st.markdown("### 📋 Detaylı İş Programı")
    df_gosterim = df_veri.copy(); df_gosterim["İlerleme %"] = df_gosterim["WBS"].map(guncel_ilerleme).fillna(0).astype(int); st.dataframe(df_gosterim, use_container_width=True)

elif secili_sekme == "CPM & Gantt":
    st.subheader("⏱️ CPM Sonuçları ve Gantt")
    t1, t2 = st.tabs(["Gantt Şeması", "Kritik Yol"])
    with t1:
        df = df_cpm.iloc[::-1].reset_index(drop=True); fig, ax = plt.subplots(figsize=(12, max(6, len(df)*0.4)))
        for i, r in df.iterrows():
            col = "#E63946" if r["Kritik Mi?"]=="EVET" else "#1976D2"; ax.barh(i, r["Süre (Gün)"], left=r["ES"], height=0.55, color=col, edgecolor="black"); ax.text(r["ES"]+r["Süre (Gün)"]/2, i, f"{r['Süre (Gün)']}g", ha='center', va='center', color='white', fontsize=7, fontweight='bold')
        ax.set_yticks(range(len(df))); ax.set_yticklabels([f"[{r['WBS Kodu']}] {r['İş Kalemi'][:28]}" for _,r in df.iterrows()], fontsize=8); ax.set_xlabel("Gün"); ax.grid(axis='x', linestyle=':', alpha=0.6); st.pyplot(fig); plt.close(fig)
    with t2: st.dataframe(df_cpm, use_container_width=True)

elif secili_sekme == "EVM & Maliyet":
    st.subheader("💰 Kazanılmış Değer Yönetimi")
    st.info("İlerleme yüzdelerini değiştirdiğinizde sistem otomatik veritabanına kaydeder.")
    edited_df = st.data_editor(evm_baslangic, column_config={"WBS Kodu": st.column_config.TextColumn("WBS", disabled=True), "İş Kalemi": st.column_config.TextColumn("İş Kalemi", disabled=True, width="large"), "Bütçe (TL)": st.column_config.NumberColumn("Bütçe (TL)", format="%,.0f", disabled=True), "İlerleme %": st.column_config.NumberColumn("İlerleme %", min_value=0, max_value=100, step=5, format="%d %%")}, hide_index=True, use_container_width=True, key="evm_editor")
    for _, row in edited_df.iterrows(): ilerleme_kaydet(row["WBS Kodu"], row["İlerleme %"])
    toplam_pv, toplam_ev, toplam_ac = 0, 0, 0
    for _, row in edited_df.iterrows(): butce = row["Bütçe (TL)"]; ilerleme = row["İlerleme %"] / 100.0; toplam_pv += butce; toplam_ev += butce * ilerleme; toplam_ac += (butce * ilerleme) * 1.10
    cpi = toplam_ev / toplam_ac if toplam_ac > 0 else 1.0; k1, k2, k3 = st.columns(3); k1.metric("Kazanılmış (EV)", f"{toplam_ev:,.0f} TL"); k2.metric("Gerçekleşen (AC)", f"{toplam_ac:,.0f} TL"); k3.metric("CPI", f"{cpi:.2f}", delta=f"{cpi-1:.2f}", delta_color="inverse")

elif secili_sekme == "Satın Alma":
    st.subheader("🛒 Satın Alma ve Tedarik")
    with st.expander("➕ Yeni Kayıt Ekle", expanded=False):
        with st.form("satin_alma_form"):
            c1, c2 = st.columns(2)
            with c1: malzeme = st.text_input("Malzeme Adı"); tedarikci = st.text_input("Tedarikçi"); tutar = st.number_input("Tutar (TL)", min_value=0.0)
            with c2: durum = st.selectbox("Durum", ["Teklif Alındı", "Sipariş Verildi", "Teslim Edildi", "İptal"]); tarih = st.date_input("Tarih")
            if st.form_submit_button("Kaydet") and malzeme:
                with conn.session as s: s.execute(text("INSERT INTO satin_alma (malzeme, tedarikci, tutar, durum, tarih) VALUES (:m, :t, :u, :d, :tr)"), {"m": malzeme, "t": tedarikci, "u": tutar, "d": durum, "tr": str(tarih)}); s.commit()
                st.success("Kayıt eklendi!"); st.rerun()
    try: st.dataframe(conn.query("SELECT id, malzeme, tedarikci, tutar, durum, tarih FROM satin_alma ORDER BY id DESC", ttl=0), use_container_width=True)
    except: st.info("Henüz kayıt yok.")

elif secili_sekme == "İşçilik":
    st.subheader("👷 İşçilik ve Puantaj")
    with st.expander("➕ Yeni Puantaj Ekle", expanded=False):
        with st.form("iscilik_form"):
            c1, c2 = st.columns(2)
            with c1: ekip = st.text_input("Ekip Adı"); is_tanimi = st.text_input("İş Tanımı"); adet = st.number_input("Kişi Sayısı", min_value=1, step=1)
            with c2: yevmiye = st.number_input("Günlük Ücret (TL)", min_value=0.0); tarih = st.date_input("Tarih")
            if st.form_submit_button("Kaydet") and ekip:
                with conn.session as s: s.execute(text("INSERT INTO iscilik (ekip, is_tanimi, adet, yevmiye, tarih) VALUES (:e, :i, :a, :y, :t)"), {"e": ekip, "i": is_tanimi, "a": adet, "y": yevmiye, "t": str(tarih)}); s.commit()
                st.success("Puantaj eklendi!"); st.rerun()
    try: st.dataframe(conn.query("SELECT id, ekip, is_tanimi, adet, yevmiye, tarih, (adet*yevmiye) as Toplam FROM iscilik ORDER BY id DESC", ttl=0), use_container_width=True)
    except: st.info("Henüz kayıt yok.")
