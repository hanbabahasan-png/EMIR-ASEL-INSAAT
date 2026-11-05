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

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

plt.rcParams['font.sans-serif'] = ['DejaVu Sans']

st.set_page_config(page_title="Ünzüle Hanbaba — İnşaat Proje Yönetimi", page_icon="🏗️", layout="wide")

# ============================================================
# VERİ ÜRETİCİ (Aktiviteler + Bütçe Verisi)
# ============================================================
@st.cache_data
def is_programi_verisi_uret():
    # WBS, İş Kalemi, Süre, Önceki, Tip, Bütçe (TL)
    veriler = [
        ["A01","Hafriyat kazı",15,"","FS", 95000],
        ["A02","Grobeton C15/20",5,"A01","FS", 61000],
        ["A03","Blokaj taş dolgu",3,"A01","FS", 7200],
        ["A04","Sıkıştırılmış dolgu",2,"A03","FS", 19000],
        ["A05","Zemin etüdü",7,"","FS", 28000],
        ["A06","Resmi işlemler (Ruhsat)",30,"","FS", 65000],
        ["B01","Radye temel betonu C30/37",20,"A02,A04","FS", 247000],
        ["B02","Radye altı membran yalıtım",5,"A02","FS", 33500],
        ["B03","Radye donatı Ø8",5,"B01","FS", 13600],
        ["B04","Radye donatı Ø10",5,"B01","FS", 21400],
        ["B05","Radye donatı Ø12",5,"B01","FS", 31200],
        ["B06","Radye donatı Ø14",5,"B01","FS", 42900],
        ["B07","Radye donatı Ø16",5,"B01","FS", 23400],
        ["B08","Radye donatı Ø20",5,"B01","FS", 7800],
        ["B09","Radye temel kalıbı",10,"B03,B04,B05,B06,B07,B08","FS", 20900],
        ["B10","Demir bağlama teli + sarf",2,"B01","FS", 9000],
        ["C01","Subasman perde betonu C30",10,"B09","FS", 43200],
        ["C02","Subasman donatı Ø12",3,"C01","FS", 35100],
        ["C03","Subasman donatı Ø14",3,"C01","FS", 23400],
        ["C04","Subasman donatı Ø16",3,"C01","FS", 9100],
        ["C05","Subasman kalıbı",5,"C02,C03,C04","FS", 39000],
        ["C06","Subasman dış yalıtım (XPS)",4,"C05","FS", 23400],
        ["C07","Drenaj borusu + çakıl",3,"C06","FS", 10000],
        ["D01","Zemin kat kolon betonu C30",15,"C07","FS", 28800],
        ["D02","Zemin kat perde betonu C30",10,"C07","FS", 21600],
        ["D03","Kolon donatı Ø16",4,"D01","FS", 35100],
        ["D04","Kolon donatı Ø20",4,"D01","FS", 27300],
        ["D05","Perde donatı Ø12",3,"D02","FS", 11700],
        ["D06","Perde donatı Ø14",3,"D02","FS", 7900],
        ["D07","Kolon kalıbı",8,"D03,D04","FS", 21400],
        ["D08","Perde kalıbı",6,"D05,D06","FS", 15600],
        # Diğer kalemler test için bütçesiz (0) bırakıldı. İstenirse eklenebilir.
    ]
    df = pd.DataFrame(veriler, columns=["WBS","İş Kalemi","Süre (Gün)","Önceki Aktivite","Bağımlılık Tipi","Bütçe (TL)"])
    
    # Boş sütunları ekle
    for col in ["Sorumlu","Plan Başlangıç","Plan Bitiş","Fiili Başlangıç","Fiili Bitiş",
                "Plan İlerleme %","Fiili İlerleme %","Ağırlık %","Ağırlıklı İlerleme %",
                "Durum","Gecikme (gün)","Kanıt / Saha Notu"]:
        df[col] = ""
        
    # Sütun sırasını düzenle (Bütçe en sona)
    df = df[["WBS","İş Kalemi","Sorumlu","Plan Başlangıç","Plan Bitiş",
             "Fiili Başlangıç","Fiili Bitiş","Plan İlerleme %","Fiili İlerleme %",
             "Ağırlık %","Ağırlıklı İlerleme %","Durum","Gecikme (gün)",
             "Kanıt / Saha Notu","Süre (Gün)","Önceki Aktivite","Bağımlılık Tipi","Bütçe (TL)"]]
    return df

# ============================================================
# CPM MOTORU
# ============================================================
@st.cache_data
def cpm_hesapla(df_raw, baslangic="2026-01-15"):
    akt = {}
    for _, row in df_raw.iterrows():
        wbs = str(row["WBS"]).strip()
        if not wbs or wbs.lower() in ('nan','none',''): continue
        try:
            dur = int(float(row["Süre (Gün)"])) if pd.notna(row["Süre (Gün)"]) else 1
            if dur <= 0: dur = 1
        except: dur = 1
        preds_raw = str(row["Önceki Aktivite"]) if pd.notna(row["Önceki Aktivite"]) else ""
        preds = [p.strip() for p in preds_raw.split(',') if p.strip()]
        dtype = str(row["Bağımlılık Tipi"]).upper() if pd.notna(row["Bağımlılık Tipi"]) else "FS"
        if dtype not in ["FS","SS","FF","SF"]: dtype = "FS"
        
        butce = float(row["Bütçe (TL)"]) if pd.notna(row["Bütçe (TL)"]) else 0.0

        akt[wbs] = {'WBS':wbs,'Task':str(row["İş Kalemi"]),'Dur':dur,
                    'Preds':preds,'Type':dtype,'Succ':[], 'Butce': butce}

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
        es[w] = max(1, e_start)
        ef[w] = es[w]+dur-1

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
        a = akt[w]; dur = a['Dur']
        slack = ls[w]-es[w]
        crit = "EVET" if slack==0 else "HAYIR"
        
        # Maliyet ağırlığı (EVM için)
        agirlik_maliyet = (a['Butce'] / toplam_butce * 100) if toplam_butce > 0 else 0
        
        rows.append({
            "WBS Kodu": w, "İş Kalemi": a['Task'], "Süre (Gün)": dur,
            "Önceki": ",".join(a['Preds']), "Bağlantı": a['Type'],
            "ES": es[w], "EF": ef[w], "LS": ls[w], "LF": lf[w],
            "Slack": slack, "Kritik Mi?": crit,
            "Plan Başlangıç": (base+timedelta(days=es[w]-1)).strftime("%d.%m.%Y"),
            "Plan Bitiş": (base+timedelta(days=ef[w]-1)).strftime("%d.%m.%Y"),
            "Ağırlık (%)": round((dur/toplam)*100,2) if toplam else 0,
            "Bütçe (TL)": a['Butce']
        })
    return pd.DataFrame(rows), proje_bitis, toplam_butce

# ============================================================
# EXCEL & PDF FONKSİYONLARI
# ============================================================
def excel_giris_olustur(df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "05_IS_PROGRAMI"
    for j, col in enumerate(df.columns, start=1):
        c = ws.cell(row=4, column=j, value=col)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1976D2")
        c.alignment = Alignment(horizontal="center")
    for i, row in df.iterrows():
        for j, val in enumerate(row, start=1):
            ws.cell(row=5+i, column=j, value=val)
    for j in range(1, len(df.columns)+1):
        ws.column_dimensions[get_column_letter(j)].width = 18
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.getvalue()

def pdf_rapor(df_cpm, proje_suresi):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    el = []
    styles = getSampleStyleSheet()
    title = ParagraphStyle('T', parent=styles['Heading1'], fontSize=14, textColor=colors.HexColor("#1976D2"), spaceAfter=8, alignment=1)
    el.append(Paragraph("<b>ÜNZÜLE HANBABA — CPM PROJE RAPORU</b>", title))
    el.append(Paragraph(f"<b>Tarih:</b> {datetime.now().strftime('%d.%m.%Y %H:%M')}", styles['Normal']))
    el.append(Spacer(1, 10))
    kritik = df_cpm[df_cpm["Kritik Mi?"]=="EVET"]
    ozet = [["Metrik","Değer"],
            ["Proje Toplam Süresi", f"{proje_suresi} gün"],
            ["Toplam Aktivite", str(len(df_cpm))],
            ["Kritik Aktivite", str(len(kritik))]]
    t = Table(ozet, colWidths=[200, 150])
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor("#1976D2")), ('TEXTCOLOR',(0,0),(-1,0),colors.white), ('GRID',(0,0),(-1,-1),0.5,colors.grey), ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold')]))
    el.append(t)
    doc.build(el); buf.seek(0)
    return buf.getvalue()

def excel_cikti(df_cpm):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df_cpm.to_excel(w, sheet_name="CPM_Detay", index=False)
        df_cpm[df_cpm["Kritik Mi?"]=="EVET"].to_excel(w, sheet_name="Kritik_Yol", index=False)
    buf.seek(0)
    return buf.getvalue()

# ============================================================
# ANA UYGULAMA
# ============================================================
st.title("🏗️ Ünzüle Hanbaba — İnşaat Proje Yönetimi Paneli")
st.caption("CPM • EVM • Gantt • Monte Carlo • Maliyet Kontrolü")

df_veri = is_programi_verisi_uret()
df_cpm, proje_suresi, toplam_butce = cpm_hesapla(df_veri)

# --- FAZ 4: SAHA İLERLEME SİMÜLASYONU (SESSION STATE) ---
if 'ilerleme' not in st.session_state:
    st.session_state.ilerleme = {row['WBS Kodu']: 0 for _, row in df_cpm.iterrows()}

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("📊 Proje Özeti")
    st.metric("Toplam Süre", f"{proje_suresi} gün")
    st.metric("Kritik Aktivite", len(df_cpm[df_cpm["Kritik Mi?"]=="EVET"]))
    st.metric("Toplam Bütçe", f"{toplam_butce:,.0f} TL")
    st.divider()
    st.subheader("📥 İndir")
    st.download_button("📄 PDF Rapor", pdf_rapor(df_cpm, proje_suresi), "CPM_Rapor.pdf", "application/pdf")
    st.download_button("📊 CPM Excel", excel_cikti(df_cpm), "CPM_Sonuc.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("📋 Giriş Excel", excel_giris_olustur(df_veri), "Villa_Is_Programi.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================
# MODÜLLER
# ============================================================
modul = st.selectbox("Modül Seç", [
    "1. İş Programı Verisi",
    "2. CPM & Kritik Yol",
    "3. Gantt Şeması",
    "4. Bütçe & Maliyet (EVM)",
    "5. Monte Carlo Simülasyonu"
])

if modul == "1. İş Programı Verisi":
    st.subheader("📋 İş Programı ve Bütçe")
    st.dataframe(df_veri, use_container_width=True)

elif modul == "2. CPM & Kritik Yol":
    st.subheader("⏱️ CPM Sonuçları")
    st.dataframe(df_cpm, use_container_width=True)

elif modul == "3. Gantt Şeması":
    st.subheader("📊 Gantt & Kritik Yol")
    df = df_cpm.iloc[::-1].reset_index(drop=True)
    fig, ax = plt.subplots(figsize=(12, max(6, len(df)*0.4)))
    for i, r in df.iterrows():
        crit = r["Kritik Mi?"]=="EVET"
        col = "#D32F2F" if crit else "#1976D2"
        ax.barh(i, r["Süre (Gün)"], left=r["ES"], height=0.55, color=col, edgecolor="black")
        ax.text(r["ES"]+r["Süre (Gün)"]/2, i, f"{r['Süre (Gün)']}g", ha='center', va='center', color='white', fontsize=7, fontweight='bold')
    ax.set_yticks(range(len(df)))
    ax.set_yticklabels([f"[{r['WBS Kodu']}] {r['İş Kalemi'][:28]}" for _,r in df.iterrows()], fontsize=8)
    ax.set_xlabel("Proje Günü")
    ax.grid(axis='x', linestyle=':', alpha=0.6)
    st.pyplot(fig); plt.close(fig)

elif modul == "4. Bütçe & Maliyet (EVM)":
    st.subheader("💰 Kazanılmış Değer Yönetimi (EVM) ve Hakediş")
    
    # 1. Simülasyon Girişi
    st.markdown("### Saha İlerlemesini Girin")
    with st.form("ilerleme_formu"):
        yeni_ilerleme = {}
        cols = st.columns(3)
        for i, row in df_cpm.iterrows():
            wbs = row["WBS Kodu"]
            with cols[i % 3]:
                yeni_ilerleme[wbs] = st.slider(f"{wbs} - {row['İş Kalemi'][:20]}", 0, 100, st.session_state.ilerleme[wbs], key=f"slider_{wbs}")
        submit = st.form_submit_button("İlerlemeyi Kaydet ve Hesapla")
        if submit:
            st.session_state.ilerleme = yeni_ilerleme

    # 2. EVM Hesaplama
    st.markdown("---")
    st.markdown("### Maliyet Performansı (EVM)")
    
    toplam_pv = 0
    toplam_ev = 0
    toplam_ac = 0
    evm_rows = []

    for _, row in df_cpm.iterrows():
        wbs = row["WBS Kodu"]
        butce = row["Bütçe (TL)"]
        ilerleme = st.session_state.ilerleme[wbs] / 100.0
        
        # Planlanan Değer (PV): Aktivitenin zamanına göre planlanan bütçe (Basit yaklaşım: tam bütçe)
        pv = butce 
        # Kazanılmış Değer (EV): Gerçekleşen ilerleme * Bütçe
        ev = butce * ilerleme
        # Gerçekleşen Maliyet (AC): Simülasyon için EV'yi %10 fazla harcadığımızı varsayalım (CPI = 0.9)
        ac = ev * 1.10 

        toplam_pv += pv
        toplam_ev += ev
        toplam_ac += ac
        
        if pv > 0 or ev > 0:
            evm_rows.append({
                "WBS": wbs, "İş Kalemi": row["İş Kalemi"], 
                "Bütçe (BAC)": butce, "İlerleme %": st.session_state.ilerleme[wbs],
                "PV": pv, "EV": ev, "AC": ac
            })

    # 3. KPI Kartları
    cpi = toplam_ev / toplam_ac if toplam_ac > 0 else 1.0
    spi = toplam_ev / toplam_pv if toplam_pv > 0 else 1.0
    eac = toplam_butce / cpi if cpi > 0 else toplam_butce

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Kazanılmış Değer (EV)", f"{toplam_ev:,.0f} TL")
    k2.metric("Gerçekleşen Maliyet (AC)", f"{toplam_ac:,.0f} TL")
    k3.metric("Maliyet Performansı (CPI)", f"{cpi:.2f}", delta=f"{cpi-1:.2f}", delta_color="inverse")
    k4.metric("Tahmini Tamamlanma (EAC)", f"{eac:,.0f} TL")

    # 4. Tablo
    st.dataframe(pd.DataFrame(evm_rows), use_container_width=True)

elif modul == "5. Monte Carlo Simülasyonu":
    st.subheader("🎲 Süre Risk Analizi")
    n = st.slider("İterasyon", 100, 2000, 500, 100)
    if st.button("Çalıştır"):
        np.random.seed(42)
        s = [sum(np.random.triangular(r["Süre (Gün)"]*0.85, r["Süre (Gün)"], r["Süre (Gün)"]*1.35) for _, r in df_cpm.iterrows()) for _ in range(n)]
        p50, p80, p90 = np.percentile(s, [50,80,90])
        c1,c2,c3 = st.columns(3)
        c1.metric("P50", f"{p50:.0f}g"); c2.metric("P80", f"{p80:.0f}g"); c3.metric("P90", f"{p90:.0f}g")
        fig, ax = plt.subplots(figsize=(10,4))
        ax.hist(s, bins=35, color="#1976D2", edgecolor="black", alpha=0.75, density=True)
        for v,c,l in [(p50,'#388E3C','P50'),(p80,'#F57C00','P80'),(p90,'#D32F2F','P90')]:
            ax.axvline(v, color=c, ls='--', lw=2, label=f"{l} ({v:.0f}g)")
        ax.set_xlabel("Gün"); ax.legend(); st.pyplot(fig); plt.close(fig)
