import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from datetime import datetime, timedelta
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from sqlalchemy import text
from PIL import Image

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

plt.rcParams['font.sans-serif'] = ['DejaVu Sans']

st.set_page_config(page_title="Emir Asel İnşaat — Proje Yönetimi", page_icon="🏗️", layout="wide")

# Logo'yu sol üst köşeye ekle
try:
    logo = Image.open("logo.png")
    st.sidebar.image(logo, use_container_width=True)
except Exception:
    st.sidebar.warning("Logo bulunamadı. Lütfen 'logo.png' dosyasını GitHub'a yükleyin.")

# ============================================================
# VERİTABANI BAĞLANTISI (NEON POSTGRESQL)
# ============================================================
try:
    conn = st.connection("neon", type="sql")
except Exception as e:
    st.error(f"Veritabanı bağlantısı kurulamadı: {e}")
    st.stop()

def init_db():
    try:
        with conn.session as s:
            s.execute(text("""
                CREATE TABLE IF NOT EXISTS ilerleme (
                    wbs_kodu TEXT PRIMARY KEY,
                    ilerleme_yuzdesi INTEGER DEFAULT 0
                )
            """))
            s.commit()
    except Exception as e:
        st.warning(f"Tablo oluşturulamadı: {e}")

def ilerleme_kaydet(wbs, yuzde):
    try:
        with conn.session as s:
            s.execute(text("""
                INSERT INTO ilerleme (wbs_kodu, ilerleme_yuzdesi) 
                VALUES (:wbs, :yuzde)
                ON CONFLICT(wbs_kodu) DO UPDATE SET ilerleme_yuzdesi = :yuzde
            """), {"wbs": wbs, "yuzde": yuzde})
            s.commit()
    except Exception as e:
        st.error(f"Kayıt hatası: {e}")

def ilerleme_yukle():
    try:
        df = conn.query("SELECT wbs_kodu, ilerleme_yuzdesi FROM ilerleme", ttl=0)
        return dict(zip(df['wbs_kodu'], df['ilerleme_yuzdesi']))
    except Exception:
        return {}

init_db()

# ============================================================
# VARSAYILAN VERİ (Aktiviteler ve Bütçe)
# ============================================================
@st.cache_data
def varsayilan_veri():
    veriler = [
        ["A01","Hafriyat kazı",15,"","FS", 95000], ["A02","Grobeton C15/20",5,"A01","FS", 61000],
        ["A03","Blokaj taş dolgu",3,"A01","FS", 7200], ["A04","Sıkıştırılmış dolgu",2,"A03","FS", 19000],
        ["A05","Zemin etüdü",7,"","FS", 28000], ["A06","Resmi işlemler (Ruhsat)",30,"","FS", 65000],
        ["B01","Radye temel betonu C30/37",20,"A02,A04","FS", 247000], ["B02","Radye altı membran yalıtım",5,"A02","FS", 33500],
        ["B03","Radye donatı Ø8",5,"B01","FS", 13600], ["B04","Radye donatı Ø10",5,"B01","FS", 21400],
        ["B05","Radye donatı Ø12",5,"B01","FS", 31200], ["B06","Radye donatı Ø14",5,"B01","FS", 42900],
        ["B07","Radye donatı Ø16",5,"B01","FS", 23400], ["B08","Radye donatı Ø20",5,"B01","FS", 7800],
        ["B09","Radye temel kalıbı",10,"B03,B04,B05,B06,B07,B08","FS", 20900], ["B10","Demir bağlama teli + sarf",2,"B01","FS", 9000],
        ["C01","Subasman perde betonu C30",10,"B09","FS", 43200], ["C02","Subasman donatı Ø12",3,"C01","FS", 35100],
        ["C03","Subasman donatı Ø14",3,"C01","FS", 23400], ["C04","Subasman donatı Ø16",3,"C01","FS", 9100],
        ["C05","Subasman kalıbı",5,"C02,C03,C04","FS", 39000], ["C06","Subasman dış yalıtım (XPS)",4,"C05","FS", 23400],
        ["C07","Drenaj borusu + çakıl",3,"C06","FS", 10000], ["D01","Zemin kat kolon betonu C30",15,"C07","FS", 28800],
        ["D02","Zemin kat perde betonu C30",10,"C07","FS", 21600], ["D03","Kolon donatı Ø16",4,"D01","FS", 35100],
        ["D04","Kolon donatı Ø20",4,"D01","FS", 27300], ["D05","Perde donatı Ø12",3,"D02","FS", 11700],
        ["D06","Perde donatı Ø14",3,"D02","FS", 7900], ["D07","Kolon kalıbı",8,"D03,D04","FS", 21400],
        ["D08","Perde kalıbı",6,"D05,D06","FS", 15600],
    ]
    df = pd.DataFrame(veriler, columns=["WBS","İş Kalemi","Süre (Gün)","Önceki Aktivite","Bağımlılık Tipi","Bütçe (TL)"])
    for col in ["Sorumlu","Plan Başlangıç","Plan Bitiş","Fiili Başlangıç","Fiili Bitiş", "Plan İlerleme %","Fiili İlerleme %","Ağırlık %","Ağırlıklı İlerleme %", "Durum","Gecikme (gün)","Kanıt / Saha Notu"]:
        df[col] = ""
    df = df[["WBS","İş Kalemi","Sorumlu","Plan Başlangıç","Plan Bitiş", "Fiili Başlangıç","Fiili Bitiş","Plan İlerleme %","Fiili İlerleme %", "Ağırlık %","Ağırlıklı İlerleme %","Durum","Gecikme (gün)", "Kanıt / Saha Notu","Süre (Gün)","Önceki Aktivite","Bağımlılık Tipi","Bütçe (TL)"]]
    return df

# ============================================================
# CPM MOTORU
# ============================================================
@st.cache_data
def cpm_hesapla(df_raw, baslangic="2026-01-15"):
    akt = {}
    for _, row in df_raw.iterrows():
        row_dict = {str(k).strip().lower(): v for k, v in row.items()}
        wbs = str(row_dict.get('wbs kodu', row_dict.get('wbs', ''))).strip()
        if not wbs or wbs.lower() in ('nan','none',''): continue
        task = str(row_dict.get('i̇ş kalemi', row_dict.get('iş kalemi', 'İsimsiz')))
        try:
            dur = int(float(row_dict.get('süre (gün)', row_dict.get('süre', 1))))
            if dur <= 0: dur = 1
        except: dur = 1
        preds_raw = str(row_dict.get('önceki aktivite', row_dict.get('önceki', '')))
        preds = [p.strip() for p in preds_raw.split(',') if p.strip() and p.strip().lower() != 'nan']
        dtype = str(row_dict.get('bağımlılık tipi', row_dict.get('tip', 'FS'))).upper()
        if dtype not in ["FS","SS","FF","SF"]: dtype = "FS"
        butce_val = row_dict.get('bütçe (tl)', row_dict.get('bütçe tutarı', row_dict.get('bütçe', 0)))
        try: butce = float(butce_val) if pd.notna(butce_val) else 0.0
        except: butce = 0.0
        akt[wbs] = {'WBS':wbs,'Task':task,'Dur':dur, 'Preds':preds,'Type':dtype,'Succ':[], 'Butce': butce}

    wbs_set = set(akt.keys())
    for w, a in akt.items():
        valid = []
        for p in a['Preds']:
            if p in wbs_set:
                valid.append(p)
                akt[p]['Succ'].append({'wbs':w,'type':a['Type']})
        a['Preds'] = valid

    in_d = {w: len(a['Preds']) for w,a in akt.items()}
    q = [w for w,d in in_d.items() if d==0]
    topo = []
    while q:
        c = q.pop(0); topo.append(c)
        for s in akt[c]['Succ']:
            in_d[s['wbs']] -= 1
            if in_d[s['wbs']]==0: q.append(s['wbs'])
            
    es, ef = {}, {}
    for w in topo:
        a = akt[w]; dur = a['Dur']
        if not a['Preds']: e_start = 1
        else:
            es_l, ef_l = [], []
            for p in a['Preds']:
                if a['Type']=="FS": es_l.append(ef[p]+1)
                elif a['Type']=="SS": es_l.append(es[p])
                elif a['Type']=="FF": ef_l.append(ef[p])
                elif a['Type']=="SF": ef_l.append(es[p])
            e_start = max(es_l) if es_l else (max(ef_l)-dur+1 if ef_l else 1)
        es[w] = max(1, e_start); ef[w] = es[w]+dur-1
        
    proje_bitis = max(ef.values()) if ef else 1
    
    ls, lf = {}, {}
    for w in reversed(topo):
        a = akt[w]; dur = a['Dur']
        if not a['Succ']: lf[w] = proje_bitis
        else:
            lf_l, ls_l = [], []
            for s in a['Succ']:
                sw = s['wbs']
                if s['type']=="FS": lf_l.append(ls[sw]-1)
                elif s['type']=="SS": ls_l.append(ls[sw])
                elif s['type']=="FF": lf_l.append(lf[sw])
                elif s['type']=="SF": ls_l.append(lf[sw])
            lf[w] = min(lf_l) if lf_l else (min(ls_l)+dur-1 if ls_l else proje_bitis)
        ls[w] = lf[w]-dur+1

    base = datetime.strptime(baslangic, "%Y-%m-%d")
    toplam = sum(a['Dur'] for a in akt.values())
    toplam_butce = sum(a['Butce'] for a in akt.values())

    rows = []
    for w in topo:
        a = akt[w]; dur = a['Dur']; slack = ls[w]-es[w]; crit = "EVET" if slack==0 else "HAYIR"
        rows.append({"WBS Kodu": w, "İş Kalemi": a['Task'], "Süre (Gün)": dur, "Önceki": ",".join(a['Preds']), "Bağlantı": a['Type'], "ES": es[w], "EF": ef[w], "LS": ls[w], "LF": lf[w], "Slack": slack, "Kritik Mi?": crit, "Plan Başlangıç": (base+timedelta(days=es[w]-1)).strftime("%d.%m.%Y"), "Plan Bitiş": (base+timedelta(days=ef[w]-1)).strftime("%d.%m.%Y"), "Ağırlık (%)": round((dur/toplam)*100,2) if toplam else 0, "Bütçe (TL)": a['Butce']})
    return pd.DataFrame(rows), proje_bitis, toplam_butce

# ============================================================
# YARDIMCI FONKSİYONLAR
# ============================================================
def pdf_rapor(df_cpm, proje_suresi):
    buf = io.BytesIO(); doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    el = []; styles = getSampleStyleSheet(); title = ParagraphStyle('T', parent=styles['Heading1'], fontSize=14, textColor=colors.HexColor("#1976D2"), spaceAfter=8, alignment=1)
    el.append(Paragraph("<b>EMİR ASEL İNŞAAT — CPM PROJE RAPORU</b>", title)); el.append(Paragraph(f"<b>Tarih:</b> {datetime.now().strftime('%d.%m.%Y %H:%M')}", styles['Normal'])); el.append(Spacer(1, 10))
    kritik = df_cpm[df_cpm["Kritik Mi?"]=="EVET"]; ozet = [["Metrik","Değer"], ["Proje Toplam Süresi", f"{proje_suresi} gün"], ["Toplam Aktivite", str(len(df_cpm))], ["Kritik Aktivite", str(len(kritik))]]
    t = Table(ozet, colWidths=[200, 150]); t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor("#1976D2")), ('TEXTCOLOR',(0,0),(-1,0),colors.white), ('GRID',(0,0),(-1,-1),0.5,colors.grey), ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold')])); el.append(t)
    doc.build(el); buf.seek(0); return buf.getvalue()

def excel_cikti(df_cpm):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df_cpm.to_excel(w, sheet_name="CPM_Detay", index=False)
        df_cpm[df_cpm["Kritik Mi?"]=="EVET"].to_excel(w, sheet_name="Kritik_Yol", index=False)
    buf.seek(0); return buf.getvalue()

# ============================================================
# ANA UYGULAMA
# ============================================================
st.title("🏗️ Emir Asel İnşaat — Proje Yönetim Paneli")
st.caption("CPM • EVM • Gantt • Monte Carlo • Nakit Akışı • Neon SQL Veritabanı")

st.sidebar.header("📂 Proje Verisi Yükle")
uploaded_excel = st.sidebar.file_uploader("Excel Yükle (WBS, Süre, Bütçe)", type=["xlsx"])

df_veri = varsayilan_veri()

if uploaded_excel is not None:
    try:
        df_yuklenen = pd.read_excel(uploaded_excel, sheet_name="05_IS_PROGRAMI", header=3)
        df_yuklenen.columns = [str(c).strip() for c in df_yuklenen.columns]
        df_veri = pd.DataFrame()
        df_veri['WBS'] = df_yuklenen.iloc[:, 0]
        df_veri['İş Kalemi'] = df_yuklenen.iloc[:, 1]
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
        for col in ["Sorumlu","Plan Başlangıç","Plan Bitiş","Fiili Başlangıç","Fiili Bitiş", "Plan İlerleme %","Fiili İlerleme %","Ağırlık %","Ağırlıklı İlerleme %", "Durum","Gecikme (gün)","Kanıt / Saha Notu"]:
            df_veri[col] = ""
        df_veri = df_veri[["WBS","İş Kalemi","Sorumlu","Plan Başlangıç","Plan Bitiş", "Fiili Başlangıç","Fiili Bitiş","Plan İlerleme %","Fiili İlerleme %", "Ağırlık %","Ağırlıklı İlerleme %","Durum","Gecikme (gün)", "Kanıt / Saha Notu","Süre (Gün)","Önceki Aktivite","Bağımlılık Tipi","Bütçe (TL)"]]
        st.sidebar.success("✅ Excel başarıyla yüklendi!")
    except Exception as e:
        st.sidebar.error(f"❌ Excel okunamadı: {e}")

df_cpm, proje_suresi, toplam_butce = cpm_hesapla(df_veri)

kayitli_ilerleme = ilerleme_yukle()
evm_baslangic = df_cpm[df_cpm["Bütçe (TL)"] > 0][["WBS Kodu", "İş Kalemi", "Bütçe (TL)"]].copy()
evm_baslangic["İlerleme %"] = evm_baslangic["WBS Kodu"].map(kayitli_ilerleme).fillna(0).astype(int)
st.session_state.evm_df = evm_baslangic

with st.sidebar:
    st.header("📊 Proje Özeti")
    st.metric("Toplam Süre", f"{proje_suresi} gün")
    st.metric("Kritik Aktivite", len(df_cpm[df_cpm["Kritik Mi?"]=="EVET"]))
    st.metric("Toplam Bütçe", f"{toplam_butce:,.0f} TL")
    st.divider()
    st.subheader("📥 İndir")
    st.download_button("📄 PDF Rapor", pdf_rapor(df_cpm, proje_suresi), "CPM_Rapor.pdf", "application/pdf")
    st.download_button("📊 CPM Excel", excel_cikti(df_cpm), "CPM_Sonuc.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

modul = st.selectbox("Modül Seç", [
    "1. İş Programı Verisi", "2. CPM & Kritik Yol", "3. Gantt Şeması", 
    "4. Bütçe & Maliyet (EVM)", "5. Monte Carlo Simülasyonu", "6. Nakit Akışı (S-Eğrisi)"
])

if modul == "1. İş Programı Verisi":
    st.subheader("📊 Proje Yönetim Paneli (Dashboard)")
    
    guncel_ilerleme = ilerleme_yukle()
    toplam_pv, toplam_ev, toplam_ac = 0, 0, 0
    en_pahali_kalemler = []
    
    for _, row in df_cpm.iterrows():
        butce = row["Bütçe (TL)"]
        if butce > 0:
            ilerleme = guncel_ilerleme.get(row["WBS Kodu"], 0) / 100.0
            toplam_pv += butce
            ev = butce * ilerleme
            toplam_ev += ev
            toplam_ac += ev * 1.10
            en_pahali_kalemler.append({"İş Kalemi": row["İş Kalemi"], "Bütçe": butce, "İlerleme": ilerleme * 100})

    c1, c2, c3, c4 = st.columns(4)
    genel_ilerleme = (toplam_ev / toplam_pv * 100) if toplam_pv > 0 else 0
    c1.metric("Genel Fiziksel İlerleme", f"%{genel_ilerleme:.1f}")
    c2.metric("Toplam Bütçe", f"{toplam_pv:,.0f} TL")
    c3.metric("Hakediş (Kazanılmış)", f"{toplam_ev:,.0f} TL")
    cpi = toplam_ev / toplam_ac if toplam_ac > 0 else 1.0
    delta_color = "inverse" if cpi < 1 else "normal"
    c4.metric("Maliyet Performansı (CPI)", f"{cpi:.2f}", delta=f"{cpi-1:.2f}", delta_color=delta_color)

    st.markdown("---")
    g1, g2 = st.columns(2)
    
    with g1:
        st.markdown("### 🏗️ Disiplinlere Göre Bütçe Dağılımı")
        df_butce_grup = df_cpm[df_cpm["Bütçe (TL)"] > 0].copy()
        df_butce_grup['Grup'] = df_butce_grup['WBS Kodu'].str[0]
        grup_isimleri = {'A': 'Hafriyat/Hazırlık', 'B': 'Temel', 'C': 'Subasman', 'D': 'Üst Yapı'}
        df_butce_grup['Disiplin'] = df_butce_grup['Grup'].map(grup_isimleri).fillna('Diğer')
        butce_ozet = df_butce_grup.groupby('Disiplin')['Bütçe (TL)'].sum().reset_index()
        
        fig1, ax1 = plt.subplots(figsize=(7, 5))
        wedges, texts, autotexts = ax1.pie(butce_ozet['Bütçe (TL)'], labels=butce_ozet['Disiplin'], autopct='%1.1f%%', startangle=90, colors=plt.cm.Set3.colors, textprops={'fontsize': 10})
        ax1.axis('equal')
        st.pyplot(fig1); plt.close(fig1)

    with g2:
        st.markdown("### 💰 En Yüksek Bütçeli 5 Kalem ve İlerlemesi")
        en_pahali_kalemler.sort(key=lambda x: x["Bütçe"], reverse=True)
        top5 = en_pahali_kalemler[:5]
        
        fig2, ax2 = plt.subplots(figsize=(7, 5))
        isimler = [k["İş Kalemi"][:20] + "..." for k in top5]
        butceler = [k["Bütçe"] for k in top5]
        bars = ax2.barh(isimler, butceler, color='#1976D2', alpha=0.7, edgecolor='black')
        
        for bar, kalem in zip(bars, top5):
            ax2.text(bar.get_width() / 2, bar.get_y() + bar.get_height()/2, f'%{kalem["İlerleme"]:.0f}', ha='center', va='center', color='white', fontweight='bold', fontsize=9)
        ax2.set_xlabel("Bütçe (TL)")
        ax2.invert_yaxis()
        st.pyplot(fig2); plt.close(fig2)

    st.markdown("---")
    st.markdown("### 📋 Detaylı İş Programı ve Saha İlerlemesi")
    df_gosterim = df_veri.copy()
    df_gosterim["İlerleme %"] = df_gosterim["WBS"].map(guncel_ilerleme).fillna(0).astype(int)
    st.dataframe(df_gosterim, use_container_width=True)
    st.info("💡 İlerleme yüzdelerini değiştirmek için '4. Bütçe & Maliyet (EVM)' sekmesini kullanın.")

elif modul == "2. CPM & Kritik Yol": 
    st.subheader("⏱️ CPM Sonuçları"); st.dataframe(df_cpm, use_container_width=True)
elif modul == "3. Gantt Şeması":
    st.subheader("📊 Gantt & Kritik Yol")
    df = df_cpm.iloc[::-1].reset_index(drop=True); fig, ax = plt.subplots(figsize=(12, max(6, len(df)*0.4)))
    for i, r in df.iterrows():
        crit = r["Kritik Mi?"]=="EVET"; col = "#D32F2F" if crit else "#1976D2"
        ax.barh(i, r["Süre (Gün)"], left=r["ES"], height=0.55, color=col, edgecolor="black")
        ax.text(r["ES"]+r["Süre (Gün)"]/2, i, f"{r['Süre (Gün)']}g", ha='center', va='center', color='white', fontsize=7, fontweight='bold')
    ax.set_yticks(range(len(df))); ax.set_yticklabels([f"[{r['WBS Kodu']}] {r['İş Kalemi'][:28]}" for _,r in df.iterrows()], fontsize=8)
    ax.set_xlabel("Proje Günü"); ax.grid(axis='x', linestyle=':', alpha=0.6); st.pyplot(fig); plt.close(fig)
elif modul == "4. Bütçe & Maliyet (EVM)":
    st.subheader("💰 Kazanılmış Değer Yönetimi (EVM) ve Hakediş")
    st.info("💡 İlerleme yüzdelerini değiştirdiğinizde sistem otomatik olarak **Neon SQL veritabanına kaydeder**.")
    
    edited_df = st.data_editor(st.session_state.evm_df, column_config={
        "WBS Kodu": st.column_config.TextColumn("WBS", disabled=True),
        "İş Kalemi": st.column_config.TextColumn("İş Kalemi", disabled=True, width="large"),
        "Bütçe (TL)": st.column_config.NumberColumn("Bütçe (TL)", format="%,.0f", disabled=True),
        "İlerleme %": st.column_config.NumberColumn("İlerleme %", min_value=0, max_value=100, step=5, format="%d %%")
    }, hide_index=True, use_container_width=True, key="evm_editor")
    
    for _, row in edited_df.iterrows():
        ilerleme_kaydet(row["WBS Kodu"], row["İlerleme %"])
    
    st.session_state.evm_df = edited_df
    st.markdown("---")
    st.markdown("### Maliyet Performansı (EVM)")
    toplam_pv, toplam_ev, toplam_ac = 0, 0, 0
    for _, row in edited_df.iterrows():
        butce = row["Bütçe (TL)"]; ilerleme = row["İlerleme %"] / 100.0
        toplam_pv += butce; toplam_ev += butce * ilerleme; toplam_ac += (butce * ilerleme) * 1.10
    cpi = toplam_ev / toplam_ac if toplam_ac > 0 else 1.0; spi = toplam_ev / toplam_pv if toplam_pv > 0 else 1.0; eac = toplam_butce / cpi if cpi > 0 else toplam_butce
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Kazanılmış Değer (EV)", f"{toplam_ev:,.0f} TL")
    k2.metric("Gerçekleşen Maliyet (AC)", f"{toplam_ac:,.0f} TL")
    k3.metric("Maliyet Performansı (CPI)", f"{cpi:.2f}", delta=f"{cpi-1:.2f}", delta_color="inverse")
    k4.metric("Tahmini Tamamlanma (EAC)", f"{eac:,.0f} TL")
elif modul == "5. Monte Carlo Simülasyonu":
    st.subheader("🎲 Süre Risk Analizi"); n = st.slider("İterasyon", 100, 2000, 500, 100)
    if st.button("Çalıştır"):
        np.random.seed(42); s = [sum(np.random.triangular(r["Süre (Gün)"]*0.85, r["Süre (Gün)"], r["Süre (Gün)"]*1.35) for _, r in df_cpm.iterrows()) for _ in range(n)]
        p50, p80, p90 = np.percentile(s, [50,80,90]); c1,c2,c3 = st.columns(3)
        c1.metric("P50", f"{p50:.0f}g"); c2.metric("P80", f"{p80:.0f}g"); c3.metric("P90", f"{p90:.0f}g")
        fig, ax = plt.subplots(figsize=(10,4)); ax.hist(s, bins=35, color="#1976D2", edgecolor="black", alpha=0.75, density=True)
        for v,c,l in [(p50,'#388E3C','P50'),(p80,'#F57C00','P80'),(p90,'#D32F2F','P90')]: ax.axvline(v, color=c, ls='--', lw=2, label=f"{l} ({v:.0f}g)")
        ax.set_xlabel("Gün"); ax.legend(); st.pyplot(fig); plt.close(fig)
elif modul == "6. Nakit Akışı (S-Eğrisi)":
    st.subheader("📈 Aylık Nakit Akışı (Cash Flow)")
    aylik_harcama = {}
    for _, row in df_cpm.iterrows():
        if row["Bütçe (TL)"] > 0:
            sure_ay = max(1, row["Süre (Gün)"] / 30.0)
            aylik_maliyet = row["Bütçe (TL)"] / sure_ay
            baslangic_ayi = (row["ES"] - 1) // 30 + 1
            bitis_ayi = (row["EF"] - 1) // 30 + 1
            for ay in range(baslangic_ayi, bitis_ayi + 1):
                aylik_harcama[ay] = aylik_harcama.get(ay, 0) + aylik_maliyet
    if aylik_harcama:
        df_nakit = pd.DataFrame(list(aylik_harcama.items()), columns=["Ay", "Aylık Harcama (TL)"]).sort_values("Ay")
        df_nakit["Kümülatif Harcama (TL)"] = df_nakit["Aylık Harcama (TL)"].cumsum()
        fig, ax1 = plt.subplots(figsize=(12, 6))
        ax1.bar(df_nakit["Ay"], df_nakit["Aylık Harcama (TL)"], color="#1976D2", alpha=0.6, label="Aylık Harcama")
        ax1.set_xlabel("Proje Ayı"); ax1.set_ylabel("Aylık Harcama (TL)", color="#1976D2")
        ax2 = ax1.twinx()
        ax2.plot(df_nakit["Ay"], df_nakit["Kümülatif Harcama (TL)"], color="#D32F2F", linewidth=3, marker='o', label="Kümülatif")
        ax2.set_ylabel("Kümülatif Harcama (TL)", color="#D32F2F")
        plt.title("Proje Nakit Akışı (S-Eğrisi)", fontsize=14, fontweight='bold')
        fig.tight_layout(); st.pyplot(fig); plt.close(fig)
        st.dataframe(df_nakit, use_container_width=True)
   
