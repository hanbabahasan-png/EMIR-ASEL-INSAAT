# -*- coding: utf-8 -*-
"""
EMİR ASEL İNŞAAT | PROJE KONTROL V9.5 (NİHAİ SÜRÜM)
- Neon PostgreSQL'e zorunlu bağlantı.
- "..." sorunu giderildi, hatalar ekrana basılıyor.
"""
from __future__ import annotations
from io import BytesIO
from datetime import datetime
import os, tempfile, sqlite3
from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st
from sqlalchemy import text

try:
    from acikpoz import parse_catalog
    ACIKPOZ_OK = True
except ImportError:
    ACIKPOZ_OK = False

APP_VERSION = "9.5"

st.set_page_config(page_title="Emir Asel | Proje Kontrol V9.5", page_icon="🏗️", layout="wide")

# ============================================================
# TEMA
# ============================================================
NAVY = "#14263D"; NAVY_2 = "#0A192F"; RED = "#C83E4D"; RED_ACCENT = "#E63946"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap');
    html, body, [class*="css"] {{ font-family: 'Montserrat', sans-serif; }}
    .stApp {{ background: linear-gradient(180deg, #F8F9FA 0%, #FFFFFF 100%); }}
    .ea-header {{ background: linear-gradient(135deg, {NAVY_2} 0%, {NAVY} 100%); padding: 26px 32px; border-radius: 16px; color: white; margin-bottom: 22px; }}
    .ea-header h1 {{ color: white !important; margin: 0; font-size: 26px; font-weight: 800; }}
    .ea-header .sub {{ color: #B8C5D6; font-size: 12px; margin-top: 6px; letter-spacing: 2px; }}
    section[data-testid="stSidebar"] {{ background: linear-gradient(180deg, {NAVY_2} 0%, {NAVY} 100%); }}
    section[data-testid="stSidebar"] * {{ color: white !important; }}
    div[data-testid="stMetric"] {{ background: white; border-radius: 12px; padding: 16px 20px; border-left: 5px solid {RED}; }}
    .section-title {{ color: {NAVY}; font-size: 18px; font-weight: 800; margin: 22px 0 10px 0; padding-left: 12px; border-left: 4px solid {RED_ACCENT}; }}
    .save-box {{ background: #E8F5E9; border-left: 5px solid #4CAF50; padding: 12px 16px; border-radius: 8px; margin: 10px 0; }}
    .neon-box {{ background: #E3F2FD; border-left: 5px solid #1976D2; padding: 12px 16px; border-radius: 8px; margin: 10px 0; }}
    .error-box {{ background: #FFEBEE; border-left: 5px solid #F44336; padding: 12px 16px; border-radius: 8px; margin: 10px 0; }}
</style>
""", unsafe_allow_html=True)

# ============================================================
# ŞEMALAR
# ============================================================
PLAN_COLS = ["wbs","poz","is_kalemi","ana_kategori","alt_kategori","birim","miktar","birim_fiyat","planlanan_tutar","baslangic","bitis","sorumlu","aciklama"]
ACTUAL_COLS = ["wbs","poz","is_kalemi","ana_kategori","birim","gerceklesen_miktar","gerceklesen_birim_fiyat","gerceklesen_tutar","tarih","tedarikci","belge_no","aciklama"]
SCHEDULE_COLS = ["wbs","aktivite","sure","oncul","bagimlilik","baslangic","bitis","fiili_baslangic","fiili_bitis","tamamlanma","agirlik","sorumlu","durum","kanit"]
RISK_COLS = ["risk_id","risk_tanimi","kategori","olasilik","etki","risk_skoru","onlem","sorumlu","durum","maliyet_etkisi"]
CASH_COLS = ["donem","planlanan","gerceklesen","net","kumulatif"]

COST_CATEGORIES = ["KABA İŞLER", "İNCE İŞLER / MEP", "DIŞ İŞLER", "GENEL GİDER / İŞÇİLİK", "Diğer"]

# ============================================================
# YARDIMCI
# ============================================================
def money(v):
    try: return f"₺ {float(v):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    except: return "₺ 0,00"

# ============================================================
# VERİTABANI BAĞLANTISI (ZORUNLU NEON)
# ============================================================
NEON_OK = False
conn = None

try:
    conn = st.connection("neon", type="sql")
    # Bağlantıyı test et
    with conn.session as s:
        s.execute(text("SELECT 1"))
    NEON_OK = True
except Exception as e:
    st.markdown(f'<div class="error-box">🚨 <b>Neon Bağlantı Hatası:</b> {e}<br>Lütfen Secrets ayarlarınızı kontrol edin.</div>', unsafe_allow_html=True)
    st.stop() # Bağlantı yoksa uygulamayı durdur

def db_load(table, cols):
    try:
        df = conn.query(f"SELECT * FROM {table}", ttl=0)
        if "id" in df.columns: df = df.drop(columns=["id"])
        for c in cols:
            if c not in df.columns: df[c] = None
        return df[cols]
    except Exception:
        return pd.DataFrame(columns=cols)

def db_save(table, df, cols):
    try:
        df_to_save = df.copy()
        for c in cols:
            if c not in df_to_save.columns: df_to_save[c] = None
        df_to_save = df_to_save[cols]

        with conn.session as s:
            s.execute(text(f"DELETE FROM {table}"))
            if not df_to_save.empty:
                for _, row in df_to_save.iterrows():
                    placeholders = ", ".join([f":{c}" for c in cols])
                    col_names = ", ".join(cols)
                    params = {c: row[c] for c in cols}
                    s.execute(text(f"INSERT INTO {table} ({col_names}) VALUES ({placeholders})"), params)
            s.commit()
        
        st.session_state.clear()
        st.rerun()
        return True
    except Exception as e:
        st.error(f"Kayıt hatası: {e}")
        return False

# ============================================================
# VERİ YÜKLE
# ============================================================
if "loaded" not in st.session_state:
    st.session_state.plan = db_load("plan", PLAN_COLS)
    st.session_state.actual = db_load("actual", ACTUAL_COLS)
    st.session_state.schedule = db_load("schedule", SCHEDULE_COLS)
    st.session_state.risk = db_load("risk", RISK_COLS)
    st.session_state.cash = db_load("cash", CASH_COLS)
    st.session_state.loaded = True

# ============================================================
# HEADER
# ============================================================
st.markdown(f'<div class="ea-header"><h1>🏗️ EMİR ASEL İNŞAAT</h1><div class="sub">PROJE KONTROL V{APP_VERSION} · {datetime.now():%d.%m.%Y %H:%M}</div></div>', unsafe_allow_html=True)

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    try: st.image("logo.png", use_container_width=True)
    except: st.markdown("### 🏗️ EMİR ASEL")
    
    MENU = ["Yönetici Dashboard","Bütçe ve Metraj","Gerçekleşen Maliyet","İş Programı / CPM","EVM Performans","Risk Yönetimi","Nakit Akışı (S-Curve)","Raporlar"]
    page = st.radio("MODÜLLER", MENU)
    st.divider()
    st.markdown('<div class="neon-box">☁️ <b>Neon PostgreSQL</b> bağlı.</div>', unsafe_allow_html=True)

# ============================================================
# SAYFA: DASHBOARD
# ============================================================
plan = st.session_state.plan.copy()
actual = st.session_state.actual.copy()
budget = float(plan["planlanan_tutar"].sum()) if not plan.empty else 0.0
actual_cost = float(actual["gerceklesen_tutar"].sum()) if not actual.empty else 0.0

if page == "Yönetici Dashboard":
    st.title("Yönetici Dashboard")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Toplam Bütçe", money(budget))
    c2.metric("Gerçekleşen", money(actual_cost))
    c3.metric("Kalan Bütçe", money(budget - actual_cost))
    c4.metric("Kayıt Sayısı", len(plan))

# ============================================================
# SAYFA: BÜTÇE
# ============================================================
elif page == "Bütçe ve Metraj":
    st.title("Bütçe ve Metraj")
    
    if ACIKPOZ_OK:
        pdf_file = st.file_uploader("ÇŞB Birim Fiyat PDF'i yükle", type=["pdf"])
        if pdf_file and st.button("PDF'i Oku ve Ekle"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                tmp.write(pdf_file.getvalue()); tmp_path = tmp.name
            try:
                result = parse_catalog(tmp_path)
                kalemler = []
                for p in result.pozes:
                    if getattr(p, "is_priced", False) and getattr(p, "fiyat", None) is not None:
                        kalemler.append({"poz": str(p.poz_no).strip(), "is_kalemi": str(p.tanim)[:200], "birim": str(p.birim), "birim_fiyat": float(p.fiyat)})
                if kalemler:
                    new_rows = pd.DataFrame(kalemler)
                    for c in PLAN_COLS:
                        if c not in new_rows.columns: new_rows[c] = None
                    st.session_state.plan = pd.concat([st.session_state.plan, new_rows[PLAN_COLS]], ignore_index=True)
                    st.success(f"{len(kalemler)} kalem eklendi! Aşağıdaki 'Kaydet' butonuna basın.")
                else: st.warning("PDF'te fiyatlı poz bulunamadı.")
            except Exception as e: st.error(f"Hata: {e}")
            finally: os.unlink(tmp_path)
    
    edited = st.data_editor(st.session_state.plan, num_rows="dynamic", use_container_width=True,
        column_config={"ana_kategori": st.column_config.SelectboxColumn(options=COST_CATEGORIES)},
        key="plan_editor")
    
    if st.button("💾 Bütçe Tablosunu Kaydet (Kalıcı)", type="primary"):
        edited = edited.copy()
        edited["planlanan_tutar"] = pd.to_numeric(edited["miktar"], errors="coerce").fillna(0) * pd.to_numeric(edited["birim_fiyat"], errors="coerce").fillna(0)
        db_save("plan", edited, PLAN_COLS)

# ============================================================
# SAYFA: GERÇEKLEŞEN
# ============================================================
elif page == "Gerçekleşen Maliyet":
    st.title("Gerçekleşen Maliyet")
    edited = st.data_editor(st.session_state.actual, num_rows="dynamic", use_container_width=True,
        column_config={"ana_kategori": st.column_config.SelectboxColumn(options=COST_CATEGORIES)},
        key="actual_editor")
    
    if st.button("💾 Gerçekleşen Tablosunu Kaydet (Kalıcı)", type="primary"):
        edited = edited.copy()
        edited["gerceklesen_tutar"] = pd.to_numeric(edited["gerceklesen_miktar"], errors="coerce").fillna(0) * pd.to_numeric(edited["gerceklesen_birim_fiyat"], errors="coerce").fillna(0)
        db_save("actual", edited, ACTUAL_COLS)

# ============================================================
# SAYFA: İŞ PROGRAMI
# ============================================================
elif page == "İş Programı / CPM":
    st.title("İş Programı / CPM")
    edited = st.data_editor(st.session_state.schedule, num_rows="dynamic", use_container_width=True, key="sched_editor")
    if st.button("💾 İş Programını Kaydet (Kalıcı)", type="primary"):
        db_save("schedule", edited, SCHEDULE_COLS)

# ============================================================
# SAYFA: RİSK
# ============================================================
elif page == "Risk Yönetimi":
    st.title("Risk Yönetimi")
    edited = st.data_editor(st.session_state.risk, num_rows="dynamic", use_container_width=True, key="risk_editor")
    if st.button("💾 Riskleri Kaydet (Kalıcı)", type="primary"):
        db_save("risk", edited, RISK_COLS)

# ============================================================
# SAYFA: NAKİT AKIŞI
# ============================================================
elif page == "Nakit Akışı (S-Curve)":
    st.title("Nakit Akışı (S-Curve)")
    edited = st.data_editor(st.session_state.cash, num_rows="dynamic", use_container_width=True, key="cash_editor")
    if st.button("💾 Nakit Akışını Kaydet (Kalıcı)", type="primary"):
        db_save("cash", edited, CASH_COLS)

# ============================================================
# SAYFA: EVM
# ============================================================
elif page == "EVM Performans":
    st.title("EVM Performans")
    st.metric("BAC (Toplam Bütçe)", money(budget))
    st.metric("AC (Gerçekleşen Maliyet)", money(actual_cost))
    if budget > 0:
        st.metric("Kalan Bütçe Oranı", f"%{(budget-actual_cost)/budget*100:.1f}")

# ============================================================
# SAYFA: RAPORLAR
# ============================================================
elif page == "Raporlar":
    st.title("Raporlar")
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        plan.to_excel(writer, sheet_name="Bütçe", index=False)
        actual.to_excel(writer, sheet_name="Gerçekleşen", index=False)
        st.session_state.schedule.to_excel(writer, sheet_name="İş Programı", index=False)
        st.session_state.risk.to_excel(writer, sheet_name="Riskler", index=False)
        st.session_state.cash.to_excel(writer, sheet_name="Nakit Akışı", index=False)
    st.download_button("📥 Excel Raporu İndir", data=output.getvalue(), 
                       file_name=f"EmirAsel_{datetime.now():%Y%m%d_%H%M}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
