# -*- coding: utf-8 -*-
"""
EMİR ASEL İNŞAAT | PROJE KONTROL V9.9
- PDF Limit 15MB'a çıkarıldı.
- Akıllı Mükerrer Kontrolü (Poz No + Fiyat Karşılaştırması).
- Türkiye Saati (UTC+3) desteği.
"""
from __future__ import annotations
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
import os, tempfile
import numpy as np
import pandas as pd
import streamlit as st
from sqlalchemy import text

try:
    from acikpoz import parse_catalog
    ACIKPOZ_OK = True
except ImportError:
    ACIKPOZ_OK = False

APP_VERSION = "9.9"
TR_TZ = ZoneInfo("Europe/Istanbul")

st.set_page_config(page_title="Emir Asel | Proje Kontrol V9.9", page_icon="🏗️", layout="wide")

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
# VERİTABANI BAĞLANTISI
# ============================================================
NEON_OK = False
conn = None

try:
    conn = st.connection("neon", type="sql")
    with conn.session as s:
        s.execute(text("SELECT 1"))
    NEON_OK = True
except Exception as e:
    st.markdown(f'<div class="error-box">🚨 <b>Neon Bağlantı Hatası:</b> {e}<br>Lütfen Secrets ayarlarınızı kontrol edin.</div>', unsafe_allow_html=True)
    st.stop()

def init_db():
    tables = {
        "plan": "id SERIAL PRIMARY KEY, wbs TEXT, poz TEXT, is_kalemi TEXT, ana_kategori TEXT, alt_kategori TEXT, birim TEXT, miktar REAL, birim_fiyat REAL, planlanan_tutar REAL, baslangic TEXT, bitis TEXT, sorumlu TEXT, aciklama TEXT",
        "actual": "id SERIAL PRIMARY KEY, wbs TEXT, poz TEXT, is_kalemi TEXT, ana_kategori TEXT, birim TEXT, gerceklesen_miktar REAL, gerceklesen_birim_fiyat REAL, gerceklesen_tutar REAL, tarih TEXT, tedarikci TEXT, belge_no TEXT, aciklama TEXT",
        "schedule": "id SERIAL PRIMARY KEY, wbs TEXT, aktivite TEXT, sure REAL, oncul TEXT, bagimlilik TEXT, baslangic TEXT, bitis TEXT, fiili_baslangic TEXT, fiili_bitis TEXT, tamamlanma REAL, agirlik REAL, sorumlu TEXT, durum TEXT, kanit TEXT",
        "risk": "id SERIAL PRIMARY KEY, risk_id TEXT, risk_tanimi TEXT, kategori TEXT, olasilik REAL, etki REAL, risk_skoru REAL, onlem TEXT, sorumlu TEXT, durum TEXT, maliyet_etkisi REAL",
        "cash": "id SERIAL PRIMARY KEY, donem TEXT, planlanan REAL, gerceklesen REAL, net REAL, kumulatif REAL",
    }
    try:
        with conn.session as s:
            for t, c in tables.items():
                s.execute(text(f"CREATE TABLE IF NOT EXISTS {t} ({c})"))
            s.commit()
    except: pass

init_db()

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
# OTURUM BAŞLATMA
# ============================================================
if "loaded" not in st.session_state:
    st.session_state.plan = db_load("plan", PLAN_COLS)
    st.session_state.actual = db_load("actual", ACTUAL_COLS)
    st.session_state.schedule = db_load("schedule", SCHEDULE_COLS)
    st.session_state.risk = db_load("risk", RISK_COLS)
    st.session_state.cash = db_load("cash", CASH_COLS)
    st.session_state.loaded = True

if "pdf_kalemler" not in st.session_state:
    st.session_state.pdf_kalemler = None

# ============================================================
# HEADER
# ============================================================
now_tr = datetime.now(TR_TZ)
st.markdown(f'<div class="ea-header"><h1>🏗️ EMİR ASEL İNŞAAT</h1><div class="sub">PROJE KONTROL V{APP_VERSION} · {now_tr.strftime("%d.%m.%Y %H:%M")}</div></div>', unsafe_allow_html=True)

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    try: st.image("logo.png", width="stretch")
    except: st.markdown("### 🏗️ EMİR ASEL")
    
    MENU = ["Yönetici Dashboard","Bütçe ve Metraj","Gerçekleşen Maliyet","İş Programı / CPM","EVM Performans","Risk Yönetimi","Nakit Akışı (S-Curve)","Raporlar"]
    page = st.radio("MODÜLLER", MENU)
    st.divider()
    st.markdown('<div class="neon-box">☁️ <b>Neon PostgreSQL</b> bağlı.</div>', unsafe_allow_html=True)

# ============================================================
# DASHBOARD
# ============================================================
plan = st.session_state.plan.copy()
actual = st.session_state.actual.copy()

if not plan.empty and "planlanan_tutar" in plan.columns:
    budget = float(pd.to_numeric(plan["planlanan_tutar"], errors="coerce").fillna(0).sum())
else:
    budget = 0.0

if not actual.empty and "gerceklesen_tutar" in actual.columns:
    actual_cost = float(pd.to_numeric(actual["gerceklesen_tutar"], errors="coerce").fillna(0).sum())
else:
    actual_cost = 0.0

if page == "Yönetici Dashboard":
    st.title("Yönetici Dashboard")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Toplam Bütçe", money(budget))
    c2.metric("Gerçekleşen", money(actual_cost))
    c3.metric("Kalan Bütçe", money(budget - actual_cost))
    c4.metric("Kayıt Sayısı", len(plan))

# ============================================================
# BÜTÇE VE METRAJ
# ============================================================
elif page == "Bütçe ve Metraj":
    st.title("Bütçe ve Metraj")
    
    if ACIKPOZ_OK:
        # Limit 15 MB'a çıkarıldı
        pdf_file = st.file_uploader("ÇŞB Birim Fiyat PDF'i yükle (Maks 15MB)", type=["pdf"])
        if pdf_file:
            if pdf_file.size > 15 * 1024 * 1024:
                st.error(f"⚠️ Dosya çok büyük! ({pdf_file.size/1024/1024:.1f} MB). Lütfen PDF'i parçalara ayırıp yükleyin.")
            else:
                if st.button("PDF'i Oku ve Karşılaştır"):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                        tmp.write(pdf_file.getvalue()); tmp_path = tmp.name
                    try:
                        result = parse_catalog(tmp_path)
                        kalemler = []
                        for p in result.pozes:
                            if getattr(p, "is_priced", False) and getattr(p, "fiyat", None) is not None:
                                kalemler.append({
                                    "poz": str(p.poz_no).strip(),
                                    "is_kalemi": str(p.tanim)[:200],
                                    "birim": str(p.birim),
                                    "birim_fiyat": float(p.fiyat)
                                })
                        
                        if kalemler:
                            mevcut_dict = {}
                            if not plan.empty:
                                for _, row in plan.iterrows():
                                    p_no = str(row["poz"]).strip()
                                    if p_no and p_no.lower() != "none":
                                        mevcut_dict[p_no] = {
                                            "fiyat": float(row["birim_fiyat"]) if pd.notna(row["birim_fiyat"]) else 0.0,
                                            "is_kalemi": str(row["is_kalemi"]).strip()
                                        }

                            for k in kalemler:
                                p_no = str(k["poz"]).strip()
                                if p_no in mevcut_dict:
                                    eski_fiyat = mevcut_dict[p_no]["fiyat"]
                                    yeni_fiyat = k["birim_fiyat"]
                                    
                                    if abs(eski_fiyat - yeni_fiyat) < 0.01:
                                        k["durum"] = "MÜKERRER (Aynı)"
                                        k["sec"] = False
                                    else:
                                        k["durum"] = "GÜNCELLEME (Fiyat Farklı)"
                                        k["sec"] = True 
                                    k["eski_fiyat"] = eski_fiyat
                                else:
                                    k["durum"] = "YENİ"
                                    k["sec"] = True
                                    k["eski_fiyat"] = None
                            
                            st.session_state.pdf_kalemler = kalemler
                            st.success("PDF okundu. Lütfen aşağıdaki seçim ekranından eklemek istediklerinizi işaretleyin.")
                            st.rerun()
                        else:
                            st.warning("PDF'te fiyatlı poz bulunamadı.")
                    except Exception as e:
                        st.error(f"PDF okuma hatası: {e}")
                    finally:
                        os.unlink(tmp_path)

    if st.session_state.pdf_kalemler is not None:
        st.markdown('<div class="section-title">📋 PDF Karşılaştırma ve Seçim Ekranı</div>', unsafe_allow_html=True)
        st.info("💡 'MÜKERRER (Aynı)' olanlar başta işaretsiz gelir. 'GÜNCELLEME' ve 'YENİ' olanlar tikli gelir.")
        
        df_kalemler = pd.DataFrame(st.session_state.pdf_kalemler)
        
        edited_df = st.data_editor(
            df_kalemler,
            column_config={
                "sec": st.column_config.CheckboxColumn("Ekle?", default=False),
                "durum": st.column_config.TextColumn("Durum", disabled=True),
                "poz": st.column_config.TextColumn("Poz No", disabled=True),
                "is_kalemi": st.column_config.TextColumn("İş Kalemi", disabled=True, width="large"),
                "birim": st.column_config.TextColumn("Birim", disabled=True),
                "eski_fiyat": st.column_config.NumberColumn("Eski Fiyat", format="%.2f", disabled=True),
                "birim_fiyat": st.column_config.NumberColumn("Yeni Fiyat", format="%.2f", disabled=True),
            },
            hide_index=True,
            width="stretch",
            key="pdf_secim_editor"
        )
        
        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("✅ Seçilenleri Tabloya Ekle / Güncelle", type="primary", width="stretch"):
                secilenler = edited_df[edited_df["sec"] == True]
                if secilenler.empty:
                    st.warning("Hiç kalem seçmediniz.")
                else:
                    yeni_plan = plan.copy()
                    for _, row in secilenler.iterrows():
                        poz = row["poz"]
                        if row["durum"] == "GÜNCELLEME (Fiyat Farklı)":
                            yeni_plan = yeni_plan[yeni_plan["poz"].astype(str).str.strip() != str(poz).strip()]
                        
                        yeni_satir = {c: None for c in PLAN_COLS}
                        yeni_satir["poz"] = poz
                        yeni_satir["is_kalemi"] = row["is_kalemi"]
                        yeni_satir["birim"] = row["birim"]
                        yeni_satir["birim_fiyat"] = row["birim_fiyat"]
                        yeni_plan = pd.concat([yeni_plan, pd.DataFrame([yeni_satir])], ignore_index=True)
                    
                    st.session_state.plan = yeni_plan
                    st.session_state.pdf_kalemler = None
                    st.success(f"{len(secilenler)} kalem başarıyla işlendi! Neon'a kaydetmek için alttaki butona basmayı unutmayın.")
                    st.rerun()
        with c2:
            if st.button("❌ İptal / Seçimi Temizle", width="stretch"):
                st.session_state.pdf_kalemler = None
                st.rerun()

    st.divider()

    st.markdown('<div class="section-title">📊 Güncel Bütçe Tablosu</div>', unsafe_allow_html=True)
    
    # "None" satırlarını temizle (Sadece boş olanları filtreler)
    temiz_plan = st.session_state.plan.copy()
    temiz_plan = temiz_plan[~temiz_plan["poz"].isna() & (temiz_plan["poz"] != "None") & (temiz_plan["poz"] != "")]

    edited_plan = st.data_editor(temiz_plan, num_rows="dynamic", width="stretch",
        column_config={"ana_kategori": st.column_config.SelectboxColumn(options=COST_CATEGORIES)},
        key="plan_editor")
    
    if st.button("💾 Bütçe Tablosunu Kaydet (Kalıcı)", type="primary"):
        edited_plan = edited_plan.copy()
        edited_plan["planlanan_tutar"] = pd.to_numeric(edited_plan["miktar"], errors="coerce").fillna(0) * pd.to_numeric(edited_plan["birim_fiyat"], errors="coerce").fillna(0)
        db_save("plan", edited_plan, PLAN_COLS)

# ============================================================
# GERÇEKLEŞEN
# ============================================================
elif page == "Gerçekleşen Maliyet":
    st.title("Gerçekleşen Maliyet")
    edited = st.data_editor(st.session_state.actual, num_rows="dynamic", width="stretch",
        column_config={"ana_kategori": st.column_config.SelectboxColumn(options=COST_CATEGORIES)},
        key="actual_editor")
    
    if st.button("💾 Gerçekleşen Tablosunu Kaydet (Kalıcı)", type="primary"):
        edited = edited.copy()
        edited["gerceklesen_tutar"] = pd.to_numeric(edited["gerceklesen_miktar"], errors="coerce").fillna(0) * pd.to_numeric(edited["gerceklesen_birim_fiyat"], errors="coerce").fillna(0)
        db_save("actual", edited, ACTUAL_COLS)

# ============================================================
# İŞ PROGRAMI
# ============================================================
elif page == "İş Programı / CPM":
    st.title("İş Programı / CPM")
    edited = st.data_editor(st.session_state.schedule, num_rows="dynamic", width="stretch", key="sched_editor")
    if st.button("💾 İş Programını Kaydet (Kalıcı)", type="primary"):
        db_save("schedule", edited, SCHEDULE_COLS)

# ============================================================
# RİSK
# ============================================================
elif page == "Risk Yönetimi":
    st.title("Risk Yönetimi")
    edited = st.data_editor(st.session_state.risk, num_rows="dynamic", width="stretch", key="risk_editor")
    if st.button("💾 Riskleri Kaydet (Kalıcı)", type="primary"):
        db_save("risk", edited, RISK_COLS)

# ============================================================
# NAKİT AKIŞI
# ============================================================
elif page == "Nakit Akışı (S-Curve)":
    st.title("Nakit Akışı (S-Curve)")
    edited = st.data_editor(st.session_state.cash, num_rows="dynamic", width="stretch", key="cash_editor")
    if st.button("💾 Nakit Akışını Kaydet (Kalıcı)", type="primary"):
        db_save("cash", edited, CASH_COLS)

# ============================================================
# EVM
# ============================================================
elif page == "EVM Performans":
    st.title("EVM Performans")
    st.metric("BAC (Toplam Bütçe)", money(budget))
    st.metric("AC (Gerçekleşen Maliyet)", money(actual_cost))
    if budget > 0:
        st.metric("Kalan Bütçe Oranı", f"%{(budget-actual_cost)/budget*100:.1f}")

# ============================================================
# RAPORLAR
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
                       file_name=f"EmirAsel_{datetime.now(TR_TZ).strftime('%Y%m%d_%H%M')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
