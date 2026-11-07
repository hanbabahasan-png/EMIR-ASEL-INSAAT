# -*- coding: utf-8 -*-
"""
EMİR ASEL İNŞAAT | ENTEGRE PROJE ERP VE KONTROL SİSTEMİ V13.2
- Streamlit 1.67+ tam uyumlu (width="stretch")
- _id sütunu data_editor'da gizli
- PDF limit 10 MB, günlük rapor 50 MB
- EVM BAC = onaylı metraj durumu
- profit_forecast sayfa filtresi
- save_dataset transaction rollback
"""
from __future__ import annotations

import os
import re
import json
import uuid
import tempfile
import hashlib
import hmac
import math
from io import BytesIO
from pathlib import Path
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from typing import Any

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from sqlalchemy import create_engine, text, inspect as sa_inspect
from sqlalchemy.engine import Engine
from sqlalchemy.exc import SQLAlchemyError
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_VERSION = "13.2"
APP_NAME = "EMİR ASEL İNŞAAT"
TR_TZ = ZoneInfo("Europe/Istanbul")
DB_FILE = Path(__file__).with_name("emir_asel_v12.db")
PDF_DIR = Path(__file__).with_name("data") / "price_library_pdfs"
PDF_DIR.mkdir(parents=True, exist_ok=True)

st.set_page_config(page_title=f"{APP_NAME} | Proje Kontrol V{APP_VERSION}", page_icon="🏗️", layout="wide", initial_sidebar_state="expanded")

# ------------------------------ TEMA ------------------------------
NAVY = "#14263D"
NAVY_2 = "#0A192F"
RED = "#C83E4D"
TEAL = "#278C82"
GOLD = "#D6A84B"
PALETTE = [NAVY, RED, TEAL, GOLD, "#6487A5", "#8D79A8", "#7D9B75", "#D17A46"]
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] {{font-family:'Montserrat',sans-serif;}}
.stApp {{background:linear-gradient(180deg,#F6F8FA 0%,#FFFFFF 100%);}}
.block-container {{padding-top:1.3rem;padding-bottom:2rem;}}
.ea-header {{background:linear-gradient(135deg,{NAVY_2} 0%,{NAVY} 100%);padding:24px 30px;border-radius:15px;color:#fff;margin-bottom:20px;}}
.ea-header h1 {{color:#fff!important;margin:0;font-size:26px;font-weight:800;}}
.ea-header .sub {{color:#C4D0DF;font-size:11px;margin-top:6px;letter-spacing:1.5px;}}
section[data-testid="stSidebar"] {{background:linear-gradient(180deg,{NAVY_2} 0%,{NAVY} 100%);}}
section[data-testid="stSidebar"] * {{color:white!important;}}
div[data-testid="stMetric"] {{background:#fff;border:1px solid #E5EAF0;border-left:4px solid {RED};padding:13px 16px;border-radius:10px;}}
.ea-note {{background:#EAF2F8;border-left:4px solid #457B9D;padding:10px 14px;border-radius:7px;margin:8px 0;}}
.ea-warn {{background:#FFF4E5;border-left:4px solid #D98E04;padding:10px 14px;border-radius:7px;margin:8px 0;}}
</style>
""", unsafe_allow_html=True)

# ------------------------------ ŞEMALAR ------------------------------
PLAN_COLS = ["_id","WBS","Poz No","İş Kalemi","Ana Kategori","Alt Kategori","Birim","Miktar","Birim Fiyat","Planlanan Tutar","Fiyat Yılı","Fiyat Kaynağı","Başlangıç","Bitiş","Sorumlu","Metraj Durumu","Açıklama"]
ACTUAL_COLS = ["_id","WBS","Poz No","İş Kalemi","Ana Kategori","Birim","Gerçekleşen Miktar","Gerçekleşen Birim Fiyat","Gerçekleşen Tutar","Tarih","Tedarikçi","Belge No","Kayıt Türü","Açıklama"]
SCHEDULE_COLS = ["_id","WBS","Aktivite","Süre (gün)","Öncül WBS","Bağımlılık Tipi","Başlangıç","Bitiş","Fiili Başlangıç","Fiili Bitiş","Tamamlanma (%)","Ağırlık (%)","Sorumlu","Durum","Kanıt / Saha Notu"]
RISK_COLS = ["_id","Risk ID","Risk Tanımı","Kategori","Olasılık (1-5)","Etki (1-5)","Risk Skoru","Önlem / Aksiyon","Sorumlu","Durum","Tahmini Maliyet Etkisi","Hedef Tarih"]
CASH_COLS = ["_id","Tarih","Tür","Açıklama","Firma / Cari","Tutar","Vade Tarihi","Durum","WBS","Belge No"]
LIB_COLS = ["_id","Poz No","İş Kalemi","Birim","Birim Fiyat","Para Birimi","Fiyat Yılı","Kaynak","Kaynak Dosya","Yayın / Güncelleme Tarihi","Aktif"]
PURCHASE_COLS = ["_id","Sipariş No","WBS","Poz No","Malzeme / Hizmet","Tedarikçi","Sipariş Tarihi","Termin Tarihi","Sipariş Miktarı","Teslim Miktarı","Birim","Birim Fiyat","Para Birimi","Sipariş Tutarı","Faturalanan Tutar","Ödenen Tutar","Fatura No","Fatura Tarihi","Durum","Not"]
LABOR_COLS = ["_id","Tarih","Ekip","İş Tanımı","Çalışan Sayısı","Gün","Kişi Günlük Ücret","Toplam İşçilik","WBS","Taşeron / Firma","Not"]
INVENTORY_COLS = ["_id","Tarih","Depo","Malzeme Kodu","Malzeme","Hareket Türü","Miktar","Birim","Birim Fiyat","Tutar","WBS","Belge No","Açıklama"]
HAKEDIS_COLS = ["_id","Hakediş No","Dönem","Tarih","Taşeron / Firma","WBS","Brüt Tutar","Kesinti","Net Hakediş","KDV Oranı","KDV Tutarı","Ödenecek Tutar","Ödenen","Durum","Açıklama"]
PROJECT_COLS = ["_id","Proje Adı","İşveren","Lokasyon","Sözleşme No","Sözleşme Tarihi","Başlangıç Tarihi","Bitiş Tarihi","Para Birimi","KDV Oranı","Risk Rezervi (%)","Proje Yöneticisi","Açıklama"]
LEGACY_CASH_COLS = ["_id","Dönem","Planlanan","Gerçekleşen","Net","Kümülatif","Kaynak"]
CONTRACT_COLS = ["_id","Sözleşme No","Sözleşme Türü","Karşı Taraf","Kapsam","Sözleşme Tarihi","Başlangıç Tarihi","Bitiş Tarihi","Orijinal Bedel","Para Birimi","Avans (%)","Teminat (%)","Ödeme Vadesi (gün)","Durum","Onay Tarihi","Onaylayan","Belge / Dosya No","Açıklama"]
CHANGE_COLS = ["_id","Değişiklik No","Sözleşme No","Para Birimi","Talep Tarihi","Talep Kaynağı","Açıklama","Gelir Etkisi","Maliyet Etkisi","Süre Etkisi (gün)","Durum","Talep Eden","Onay Tarihi","Onaylayan","Belge / RFI No"]
EMPLOYER_HAKEDIS_COLS = ["_id","Hakediş No","Sözleşme No","Dönem","Tarih","WBS","Brüt İmalat Tutarı","Avans Mahsubu","Teminat Kesintisi","Diğer Kesintiler","Kesinti Toplamı","KDV Oranı","KDV Tutarı","Net Tahakkuk","Hakediş Toplamı","Fatura No","Vade Tarihi","Tahsil Edilen","Kalan Alacak","Durum","Onay Tarihi","Onaylayan","Açıklama"]
DAILY_REPORT_COLS = ["_id","Rapor Tarihi","Hava Durumu","Sıcaklık (°C)","Şantiye Sorumlusu","Toplam Personel","Alt Yükleniciler / Ekipler","Çalışan Makineler","Günlük İmalat Özeti","Beton (m³)","Donatı (ton)","Kalıp (m²)","Kazı / Dolgu (m³)","Gelen Malzemeler","Gecikme / Engel","İSG / Olay Bilgisi","Yarınki Plan","Açıklama","Ek Dosyalar","Raporlayan"]
APPROVAL_COLS = ["_id","Talep Tarihi","Veri Seti","Modül","Kayıt ID","Talep Başlığı","Tutar","Para Birimi","Talep Eden","Onay Rolü","Durum","Karar Tarihi","Karar Veren","Karar Notu"]
COST_CATEGORIES = [
    "Mimarlık / Proje Gideri", "Yapı Denetim Gideri", "Yapı Analiz Laboratuvar Gideri",
    "Hafriyat", "Grobeton ve Beton", "Toplam Kalıp + Demir İşçiliği",
    "Demir / Donatı Malzemesi", "Elektrik + Su İşçiliği", "Elektrik Malzemeleri",
    "Mekanik Tesisat", "Duvar ve Sıva", "Yalıtım", "Şap ve Zemin", "Seramik ve Kaplama",
    "Kapı / Pencere / Doğrama", "Cephe", "Çatı", "Boya", "Peyzaj ve Çevre Düzenleme",
    "Nakliye / Lojistik", "Şantiye Genel Giderleri", "İş Makinesi / Ekipman", "İSG / Çevre", "Diğer"
]
DATASETS = {
    "plan": PLAN_COLS, "actual": ACTUAL_COLS, "schedule": SCHEDULE_COLS,
    "risk": RISK_COLS, "cash": CASH_COLS, "purchase": PURCHASE_COLS,
    "labor": LABOR_COLS, "inventory": INVENTORY_COLS, "hakediş": HAKEDIS_COLS,
    "library": LIB_COLS, "project": PROJECT_COLS, "legacy_cash": LEGACY_CASH_COLS,
    "contracts": CONTRACT_COLS, "changes": CHANGE_COLS, "employer_hakediş": EMPLOYER_HAKEDIS_COLS,
    "daily_reports": DAILY_REPORT_COLS, "approvals": APPROVAL_COLS,
}
DATE_FIELDS = {
    "plan": ["Başlangıç", "Bitiş"], "actual": ["Tarih"],
    "schedule": ["Başlangıç", "Bitiş", "Fiili Başlangıç", "Fiili Bitiş"],
    "risk": ["Hedef Tarih"], "cash": ["Tarih", "Vade Tarihi"],
    "purchase": ["Sipariş Tarihi", "Termin Tarihi", "Fatura Tarihi"], "labor": ["Tarih"],
    "inventory": ["Tarih"], "hakediş": ["Tarih"],
    "library": ["Yayın / Güncelleme Tarihi"],
    "project": ["Sözleşme Tarihi", "Başlangıç Tarihi", "Bitiş Tarihi"],
    "contracts": ["Sözleşme Tarihi", "Başlangıç Tarihi", "Bitiş Tarihi", "Onay Tarihi"],
    "changes": ["Talep Tarihi", "Onay Tarihi"],
    "employer_hakediş": ["Tarih", "Vade Tarihi", "Onay Tarihi"],
    "daily_reports": ["Rapor Tarihi"],
    "approvals": ["Talep Tarihi", "Karar Tarihi"],
}

# ------------------------------ YARDIMCILAR ------------------------------
def now_tr() -> datetime:
    return datetime.now(TR_TZ)

def safe_num(value: Any, default: float = 0.0) -> float:
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass
    if isinstance(value, (int, float, np.number)):
        return float(value)
    s = str(value).strip().replace("₺", "").replace("$", "").replace("€", "").replace(" ", "")
    if not s:
        return default
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except (TypeError, ValueError):
        return default

def currency_or_default(value: Any, default: str = "₺") -> str:
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass
    value = str(value).strip()
    return default if value.casefold() in ("", "nan", "none") else value

def money(value: Any, currency: str = "₺") -> str:
    try:
        if value is None or pd.isna(value):
            return "—"
        return f"{currency} {float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return f"{currency} 0,00"

def pct(value: Any) -> str:
    try:
        return f"%{float(value):.1f}".replace(".", ",")
    except Exception:
        return "—"

def uid() -> str:
    return uuid.uuid4().hex

def json_safe_value(obj: Any):
    if obj is None:
        return None
    if isinstance(obj, (datetime, date, pd.Timestamp)):
        return None if pd.isna(obj) else obj.isoformat()
    if isinstance(obj, dict):
        return {str(k): json_safe_value(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [json_safe_value(v) for v in obj]
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating, float)):
        value = float(obj)
        return None if math.isnan(value) or math.isinf(value) else value
    if isinstance(obj, (np.bool_, bool)):
        return bool(obj)
    try:
        missing = pd.isna(obj)
        if isinstance(missing, (bool, np.bool_)) and missing:
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(obj, (str, int)):
        return obj
    return str(obj)

def to_payload(row: dict) -> str:
    return json.dumps(json_safe_value(row), ensure_ascii=False, allow_nan=False)

def ensure_frame(df: pd.DataFrame | None, columns: list[str]) -> pd.DataFrame:
    if df is None:
        df = pd.DataFrame()
    out = df.copy()
    for col in columns:
        if col not in out.columns:
            out[col] = None
    out = out[columns]
    if "_id" in columns:
        out["_id"] = out["_id"].apply(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() not in ("", "nan", "None") else uid())
        duplicate_mask = out["_id"].duplicated(keep="first")
        if duplicate_mask.any():
            out.loc[duplicate_mask, "_id"] = [uid() for _ in range(int(duplicate_mask.sum()))]
    return out

def auto_categorize(description: str) -> str:
    t = str(description).casefold()
    if any(k in t for k in ["mimari proje", "statik proje", "mekanik proje", "elektrik proje", "mimarlık"]): return "Mimarlık / Proje Gideri"
    if "denetim" in t: return "Yapı Denetim Gideri"
    if any(k in t for k in ["laboratuvar", "beton deney", "zemin deney"]): return "Yapı Analiz Laboratuvar Gideri"
    if any(k in t for k in ["kalıp işçilik", "kalip iscilik", "demir işçilik", "demir iscilik", "kalıp ve demir"]): return "Toplam Kalıp + Demir İşçiliği"
    if any(k in t for k in ["elektrik işçilik", "su tesisat işçilik", "tesisat işçilik", "elektrik ve su işçilik"]): return "Elektrik + Su İşçiliği"
    if any(k in t for k in ["demir", "donatı", "donati", "nervürlü"]): return "Demir / Donatı Malzemesi"
    if any(k in t for k in ["beton", "grobeton", "radye", "temel"]): return "Grobeton ve Beton"
    if any(k in t for k in ["hafriyat", "kazı", "kazi", "dolgu", "iksa"]): return "Hafriyat"
    if any(k in t for k in ["elektrik kablo", "pano", "armatür", "priz", "anahtar"]): return "Elektrik Malzemeleri"
    if any(k in t for k in ["mekanik", "havalandırma", "hvac", "yangın tesisat", "sıhhi tesisat"]): return "Mekanik Tesisat"
    if any(k in t for k in ["seramik", "sıva", "boya", "kapı", "pencere", "yalıtım", "izolasyon", "şap", "sap"]):
        if "yalıtım" in t or "izolasyon" in t: return "Yalıtım"
        if "seramik" in t: return "Seramik ve Kaplama"
        if "boya" in t: return "Boya"
        if "kapı" in t or "pencere" in t: return "Kapı / Pencere / Doğrama"
        if "şap" in t or "sap" in t: return "Şap ve Zemin"
        return "Duvar ve Sıva"
    if any(k in t for k in ["peyzaj", "çevre", "drenaj", "istinat", "bahçe", "yol"]): return "Peyzaj ve Çevre Düzenleme"
    if any(k in t for k in ["nakliye", "lojistik", "taşıma"]): return "Nakliye / Lojistik"
    if any(k in t for k in ["iş makinesi", "ekskavatör", "vinç", "manitou"]): return "İş Makinesi / Ekipman"
    return "Diğer"

def norm_col(value: Any) -> str:
    tr = str.maketrans("çğıöşüİ", "cgiosui")
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower().translate(tr))

def find_column(df: pd.DataFrame, candidates: list[str]):
    lookup = {norm_col(c): c for c in df.columns}
    for candidate in candidates:
        key = norm_col(candidate)
        if key in lookup:
            return lookup[key]
    for candidate in candidates:
        key = norm_col(candidate)
        for normalized, original in lookup.items():
            if key and (key in normalized or normalized in key):
                return original
    return None

def calc_plan_amount(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Miktar"] = out["Miktar"].apply(safe_num)
    out["Birim Fiyat"] = out["Birim Fiyat"].apply(safe_num)
    out["Planlanan Tutar"] = out["Miktar"] * out["Birim Fiyat"]
    return out

def calc_actual_amount(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Gerçekleşen Miktar"] = out["Gerçekleşen Miktar"].apply(safe_num)
    out["Gerçekleşen Birim Fiyat"] = out["Gerçekleşen Birim Fiyat"].apply(safe_num)
    out["Gerçekleşen Tutar"] = out["Gerçekleşen Miktar"] * out["Gerçekleşen Birim Fiyat"]
    return out

def summarize_costs(plan: pd.DataFrame, actual: pd.DataFrame) -> pd.DataFrame:
    p = plan.groupby("Ana Kategori", dropna=False)["Planlanan Tutar"].sum().rename("Bütçe") if not plan.empty else pd.Series(dtype=float, name="Bütçe")
    a = actual.groupby("Ana Kategori", dropna=False)["Gerçekleşen Tutar"].sum().rename("Gerçekleşen") if not actual.empty else pd.Series(dtype=float, name="Gerçekleşen")
    out = pd.concat([p, a], axis=1).fillna(0).reset_index()
    if out.empty:
        return pd.DataFrame(columns=["Ana Kategori", "Bütçe", "Gerçekleşen", "Sapma", "Sapma (%)"])
    out["Sapma"] = out["Gerçekleşen"] - out["Bütçe"]
    out["Sapma (%)"] = np.where(out["Bütçe"] != 0, out["Sapma"] / out["Bütçe"] * 100, np.nan)
    return out.sort_values("Bütçe", ascending=False)

def weighted_progress(schedule: pd.DataFrame) -> float:
    if schedule.empty:
        return 0.0
    progress = pd.to_numeric(schedule["Tamamlanma (%)"], errors="coerce").fillna(0).clip(0, 100)
    weights = pd.to_numeric(schedule["Ağırlık (%)"], errors="coerce").fillna(0).clip(lower=0)
    if weights.sum() <= 0:
        weights = pd.to_numeric(schedule["Süre (gün)"], errors="coerce").fillna(0).clip(lower=0)
    return float(np.average(progress, weights=weights)) if weights.sum() > 0 else float(progress.mean())

# ------------------------------ CPM ------------------------------
def parse_predecessors(raw: Any, default_rel: str = "FS") -> list[tuple[str, str, float]]:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return []
    result = []
    for token in re.split(r"[,;]", str(raw)):
        token = token.strip()
        if not token or token.casefold() == "nan":
            continue
        parts = token.split(":")
        parent = parts[0].strip()
        rel = parts[1].strip().upper() if len(parts) > 1 and parts[1].strip() else default_rel
        lag = safe_num(parts[2], 0) if len(parts) > 2 else 0.0
        if rel not in {"FS", "SS", "FF", "SF"}:
            rel = default_rel if default_rel in {"FS", "SS", "FF", "SF"} else "FS"
        result.append((parent, rel, lag))
    return result

def calculate_cpm(schedule: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    if schedule.empty:
        return pd.DataFrame(), []
    df = ensure_frame(schedule, SCHEDULE_COLS).copy()
    df["WBS"] = df["WBS"].fillna("").astype(str).str.strip()
    df["Süre (gün)"] = pd.to_numeric(df["Süre (gün)"], errors="coerce").fillna(0).clip(lower=0)
    warnings: list[str] = []
    if (df["WBS"] == "").any():
        warnings.append("WBS boş olan aktiviteler var; bu satırlar CPM hesabına alınmadı.")
        df = df[df["WBS"] != ""].copy()
    if df["WBS"].duplicated().any():
        warnings.append("Tekrarlanan WBS kodu var. Her aktivite için benzersiz WBS kullanın.")
        return df.assign(**{"Erken Başlangıç": np.nan, "Erken Bitiş": np.nan, "Geç Başlangıç": np.nan, "Geç Bitiş": np.nan, "Bolluk (gün)": np.nan, "Kritik mi?": "WBS tekrarlı"}), warnings
    duration = dict(zip(df["WBS"], df["Süre (gün)"].astype(float)))
    preds: dict[str, list[tuple[str, str, float]]] = {}
    succ: dict[str, list[tuple[str, str, float]]] = {w: [] for w in duration}
    indegree = {w: 0 for w in duration}
    for _, row in df.iterrows():
        task = row["WBS"]
        rel_default = str(row.get("Bağımlılık Tipi") or "FS").strip().upper()
        if rel_default not in {"FS", "SS", "FF", "SF"}:
            rel_default = "FS"
        edges = parse_predecessors(row.get("Öncül WBS"), rel_default)
        preds[task] = edges
        for parent, rel, lag in edges:
            if parent not in duration:
                warnings.append(f"{task}: öncül WBS '{parent}' bulunamadı.")
                continue
            succ[parent].append((task, rel, lag))
            indegree[task] += 1
    queue = [w for w, d in indegree.items() if d == 0]
    order: list[str] = []
    while queue:
        node = queue.pop(0)
        order.append(node)
        for nxt, _, _ in succ[node]:
            indegree[nxt] -= 1
            if indegree[nxt] == 0:
                queue.append(nxt)
    if len(order) != len(duration):
        warnings.append("Döngüsel öncül ilişkisi tespit edildi. CPM hesaplanamadı.")
        return df.assign(**{"Erken Başlangıç": np.nan, "Erken Bitiş": np.nan, "Geç Başlangıç": np.nan, "Geç Bitiş": np.nan, "Bolluk (gün)": np.nan, "Kritik mi?": "Döngü var"}), warnings
    es: dict[str, float] = {}
    ef: dict[str, float] = {}
    for task in order:
        d = duration[task]
        constraints = [1.0]
        for parent, rel, lag in preds[task]:
            if parent not in ef:
                continue
            if rel == "FS": constraints.append(ef[parent] + lag + 1)
            elif rel == "SS": constraints.append(es[parent] + lag)
            elif rel == "FF": constraints.append(ef[parent] + lag - d + 1)
            elif rel == "SF": constraints.append(es[parent] + lag - d + 1)
        es[task] = max(constraints)
        ef[task] = es[task] + max(d, 0) - 1
    project_finish = max(ef.values(), default=1.0)
    ls: dict[str, float] = {}
    lf: dict[str, float] = {}
    for task in reversed(order):
        d = duration[task]
        constraints = []
        for nxt, rel, lag in succ[task]:
            if nxt not in ls:
                continue
            if rel == "FS": constraints.append(ls[nxt] - lag - 1)
            elif rel == "SS": constraints.append(ls[nxt] - lag + d - 1)
            elif rel == "FF": constraints.append(lf[nxt] - lag)
            elif rel == "SF": constraints.append(lf[nxt] - lag + d - 1)
        lf[task] = min(constraints) if constraints else project_finish
        ls[task] = lf[task] - max(d, 0) + 1
    df["Erken Başlangıç"] = df["WBS"].map(es)
    df["Erken Bitiş"] = df["WBS"].map(ef)
    df["Geç Başlangıç"] = df["WBS"].map(ls)
    df["Geç Bitiş"] = df["WBS"].map(lf)
    df["Bolluk (gün)"] = df["WBS"].map(ls) - df["WBS"].map(es)
    df["Kritik mi?"] = np.where(df["Bolluk (gün)"].abs() < 1e-8, "EVET", "Hayır")
    return df, warnings

# ------------------------------ EVM (Düzeltilmiş BAC) ------------------------------
def calculate_evm(plan: pd.DataFrame, actual: pd.DataFrame, schedule: pd.DataFrame, report_date: date) -> dict[str, float]:
    # BAC = onaylı metraj durumundaki satırlar (yoksa tüm bütçe)
    if not plan.empty and "Metraj Durumu" in plan.columns:
        approved_plan = plan[plan["Metraj Durumu"].astype(str).str.strip() == "Onaylandı"]
        if approved_plan.empty:
            approved_plan = plan
        bac = float(pd.to_numeric(approved_plan.get("Planlanan Tutar", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    else:
        bac = float(pd.to_numeric(plan.get("Planlanan Tutar", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not plan.empty else 0.0
    ac = float(pd.to_numeric(actual.get("Gerçekleşen Tutar", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not actual.empty else 0.0
    pv = 0.0
    ev = 0.0
    if not plan.empty and not schedule.empty:
        sch = schedule.copy()
        sch["WBS"] = sch["WBS"].fillna("").astype(str).str.strip()
        sch["Tamamlanma (%)"] = pd.to_numeric(sch["Tamamlanma (%)"], errors="coerce").fillna(0).clip(0, 100)
        sch["Süre (gün)"] = pd.to_numeric(sch["Süre (gün)"], errors="coerce").fillna(0).clip(lower=0)
        sch["Ağırlık (%)"] = pd.to_numeric(sch["Ağırlık (%)"], errors="coerce").fillna(0).clip(lower=0)
        sch["Başlangıç"] = pd.to_datetime(sch["Başlangıç"], errors="coerce", format="mixed", dayfirst=True)
        sch["Bitiş"] = pd.to_datetime(sch["Bitiş"], errors="coerce", format="mixed", dayfirst=True)
        progress_map: dict[str, float] = {}
        planned_map: dict[str, float] = {}
        for wbs, grp in sch.groupby("WBS"):
            weights = grp["Ağırlık (%)"].copy()
            if weights.sum() <= 0:
                weights = grp["Süre (gün)"].replace(0, np.nan).fillna(0)
            physical = float(np.average(grp["Tamamlanma (%)"], weights=weights)) if weights.sum() > 0 else float(grp["Tamamlanma (%)"].mean())
            progress_map[wbs] = physical / 100.0
            fractions, fraction_weights = [], []
            for _, task in grp.iterrows():
                start, end = task["Başlangıç"], task["Bitiş"]
                if pd.isna(start) or pd.isna(end) or end < start:
                    continue
                rd = pd.Timestamp(report_date)
                if rd < start: fraction = 0.0
                elif rd >= end: fraction = 1.0
                else:
                    span = max((end - start).days, 1)
                    fraction = min(max((rd - start).days / span, 0.0), 1.0)
                fractions.append(fraction)
                task_weight = float(task["Ağırlık (%)"]) if float(task["Ağırlık (%)"]) > 0 else float(task["Süre (gün)"])
                fraction_weights.append(task_weight if task_weight > 0 else 1.0)
            if fractions:
                planned_map[wbs] = float(np.average(fractions, weights=fraction_weights))
        for _, row in plan.iterrows():
            wbs = str(row.get("WBS") or "").strip()
            amount = safe_num(row.get("Planlanan Tutar"))
            ev += amount * progress_map.get(wbs, 0.0)
            pv += amount * planned_map.get(wbs, 0.0)
    cpi = ev / ac if ac > 0 else np.nan
    spi = ev / pv if pv > 0 else np.nan
    cv, sv = ev - ac, ev - pv
    eac = bac / cpi if pd.notna(cpi) and cpi > 0 else np.nan
    return {"BAC": bac, "PV": pv, "EV": ev, "AC": ac, "CV": cv, "SV": sv,
            "CPI": cpi, "SPI": spi, "EAC": eac, "VAC": bac - eac if pd.notna(eac) else np.nan}

# ------------------------------ VERİTABANI (init_db try/except) ------------------------------
def resolve_database_url() -> str:
    url = os.getenv("DATABASE_URL", "").strip()
    if not url:
        try:
            url = str(st.secrets.get("DATABASE_URL", "")).strip()
        except Exception:
            url = ""
    if not url:
        return f"sqlite:///{DB_FILE.as_posix()}"
    if url.startswith("postgres://"):
        url = "postgresql+psycopg://" + url[len("postgres://"):]
    elif url.startswith("postgresql://"):
        url = "postgresql+psycopg://" + url[len("postgresql://"):]
    return url

@st.cache_resource(show_spinner=False)
def get_engine() -> Engine:
    url = resolve_database_url()
    connect_args = {"check_same_thread": False, "timeout": 30} if url.startswith("sqlite") else {}
    return create_engine(url, pool_pre_ping=True, future=True, connect_args=connect_args)

ENGINE = get_engine()
DB_KIND = ENGINE.dialect.name

def init_db() -> None:
    ddl = """
    CREATE TABLE IF NOT EXISTS ea_records (
        project_key TEXT NOT NULL,
        dataset TEXT NOT NULL,
        record_id TEXT NOT NULL,
        payload TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        PRIMARY KEY (project_key, dataset, record_id)
    )
    """
    with ENGINE.begin() as con:
        con.execute(text(ddl))
        try:
            con.execute(text("CREATE INDEX IF NOT EXISTS ix_ea_records_project_dataset ON ea_records(project_key, dataset)"))
        except Exception:
            pass

def load_dataset(project_key: str, dataset: str, columns: list[str]) -> pd.DataFrame:
    with ENGINE.connect() as con:
        rows = con.execute(text("SELECT payload FROM ea_records WHERE project_key=:p AND dataset=:d ORDER BY updated_at, record_id"), {"p": project_key, "d": dataset}).fetchall()
    records = []
    for row in rows:
        try:
            records.append(json.loads(row[0]))
        except (ValueError, TypeError):
            continue
    frame = ensure_frame(pd.DataFrame(records), columns)
    for col in DATE_FIELDS.get(dataset, []):
        if col in frame.columns:
            parsed = pd.to_datetime(frame[col], errors="coerce", format="mixed", dayfirst=True)
            frame[col] = parsed.dt.date
    return frame

def normalize_dataset_dates(frame: pd.DataFrame, dataset: str) -> pd.DataFrame:
    out = frame.copy()
    for col in DATE_FIELDS.get(dataset, []):
        if col in out.columns:
            parsed = pd.to_datetime(out[col], errors="coerce", format="mixed", dayfirst=True)
            out[col] = parsed.dt.date
    return out

def audit_log(project_key: str, action: str, dataset: str, count: int, detail: str = "") -> None:
    record = {"_id": uid(), "Tarih": now_tr().isoformat(timespec="seconds"),
              "Kullanıcı": (st.session_state.get("auth_user", {}) or {}).get("username", os.getenv("USER", "uygulama")),
              "İşlem": action, "Veri Seti": dataset, "Kayıt Sayısı": int(count), "Detay": detail}
    payload = to_payload(record)
    with ENGINE.begin() as con:
        con.execute(text("INSERT INTO ea_records(project_key,dataset,record_id,payload,updated_at) VALUES(:p,'audit',:r,:v,:t)"),
                    {"p": project_key, "r": record["_id"], "v": payload, "t": record["Tarih"]})

def save_dataset(project_key: str, dataset: str, df: pd.DataFrame, columns: list[str], make_audit: bool = True) -> None:
    """Dataset bazında transaction + record_id upsert. Hata olursa rollback."""
    frame = ensure_frame(df, columns)
    now_iso = now_tr().isoformat(timespec="seconds")
    statement = text("""
        INSERT INTO ea_records(project_key,dataset,record_id,payload,updated_at)
        VALUES(:p,:d,:r,:v,:t)
        ON CONFLICT(project_key,dataset,record_id)
        DO UPDATE SET payload=excluded.payload, updated_at=excluded.updated_at
    """)
    try:
        with ENGINE.begin() as con:
            current = con.execute(text("SELECT record_id FROM ea_records WHERE project_key=:p AND dataset=:d"), {"p": project_key, "d": dataset}).fetchall()
            current_ids = {str(row[0]) for row in current}
            keep_ids = set(frame["_id"].astype(str).tolist()) if "_id" in frame.columns else set()
            for stale_id in current_ids - keep_ids:
                con.execute(text("DELETE FROM ea_records WHERE project_key=:p AND dataset=:d AND record_id=:r"),
                            {"p": project_key, "d": dataset, "r": stale_id})
            for _, row in frame.iterrows():
                con.execute(statement, {"p": project_key, "d": dataset, "r": str(row["_id"]), "v": to_payload(row.to_dict()), "t": now_iso})
    except SQLAlchemyError as e:
        st.error(f"Kayıt sırasında veritabanı hatası: {e}")
        return
    if make_audit:
        audit_log(project_key, "KAYDET", dataset, len(frame), "Dataset bazlı atomik kayıt")

def make_snapshot(project_key: str, datasets: dict[str, pd.DataFrame]) -> str:
    snap_id = uid()
    snapshot = {name: df.to_dict(orient="records") for name, df in datasets.items()}
    record = {"_id": snap_id, "Tarih": now_tr().isoformat(timespec="seconds"), "Veri": snapshot}
    with ENGINE.begin() as con:
        con.execute(text("INSERT INTO ea_records(project_key,dataset,record_id,payload,updated_at) VALUES(:p,'snapshots',:r,:v,:t)"),
                    {"p": project_key, "r": snap_id, "v": to_payload(record), "t": record["Tarih"]})
    return snap_id

def migrate_project_records(old_key: str, new_key: str) -> None:
    if old_key == new_key:
        return
    with ENGINE.begin() as con:
        rows = con.execute(text("SELECT dataset, record_id FROM ea_records WHERE project_key=:old"), {"old": old_key}).fetchall()
        for dataset, record_id in rows:
            exists = con.execute(text("SELECT 1 FROM ea_records WHERE project_key=:new AND dataset=:d AND record_id=:r"),
                                 {"new": new_key, "d": dataset, "r": record_id}).first()
            if exists:
                continue
            con.execute(text("UPDATE ea_records SET project_key=:new WHERE project_key=:old AND dataset=:d AND record_id=:r"),
                        {"new": new_key, "old": old_key, "d": dataset, "r": record_id})

try:
    init_db()
    DB_ERROR = ""
except Exception as exc:
    DB_ERROR = str(exc)

# ------------------------------ İÇE AKTARMA ------------------------------
def map_imported_sheet(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    mappings = {
        "plan": {
            "WBS": ["WBS", "WBS Kodu", "Aktivite Kodu"], "Poz No": ["Poz No", "Poz", "İmalat Kodu"],
            "İş Kalemi": ["İş Kalemi", "Açıklama", "İmalat", "Kalem"], "Ana Kategori": ["Ana Kategori", "Kategori", "İş Grubu"],
            "Alt Kategori": ["Alt Kategori", "Alt Grup"], "Birim": ["Birim", "Ölçü Birimi"],
            "Miktar": ["Miktar", "Metraj", "Adet"], "Birim Fiyat": ["Birim Fiyat", "BF", "Fiyat"],
            "Fiyat Yılı": ["Fiyat Yılı", "Yıl"], "Fiyat Kaynağı": ["Fiyat Kaynağı", "Kaynak"],
            "Başlangıç": ["Başlangıç", "Başlangıç Tarihi"], "Bitiş": ["Bitiş", "Bitiş Tarihi"],
            "Sorumlu": ["Sorumlu", "Ekip", "Firma"], "Metraj Durumu": ["Metraj Durumu", "Durum", "Onay"], "Açıklama": ["Açıklama", "Not"]},
        "actual": {
            "WBS": ["WBS", "WBS Kodu"], "Poz No": ["Poz No", "Poz"], "İş Kalemi": ["İş Kalemi", "Açıklama", "İmalat"],
            "Ana Kategori": ["Ana Kategori", "Kategori"], "Birim": ["Birim"], "Gerçekleşen Miktar": ["Gerçekleşen Miktar", "Miktar", "Metraj"],
            "Gerçekleşen Birim Fiyat": ["Gerçekleşen Birim Fiyat", "Birim Fiyat", "BF"], "Tarih": ["Tarih", "Fatura Tarihi"],
            "Tedarikçi": ["Tedarikçi", "Firma"], "Belge No": ["Belge No", "Fatura No"], "Kayıt Türü": ["Kayıt Türü", "Tür"], "Açıklama": ["Açıklama", "Not"]},
        "schedule": {
            "WBS": ["WBS", "WBS Kodu"], "Aktivite": ["Aktivite", "İş Kalemi", "Görev"], "Süre (gün)": ["Süre (gün)", "Süre", "Duration"],
            "Öncül WBS": ["Öncül WBS", "Öncül", "Predecessor"], "Bağımlılık Tipi": ["Bağımlılık Tipi", "İlişki Tipi", "Tip"],
            "Başlangıç": ["Başlangıç", "Plan Başlangıç"], "Bitiş": ["Bitiş", "Plan Bitiş"],
            "Fiili Başlangıç": ["Fiili Başlangıç", "Gerçek Başlangıç"], "Fiili Bitiş": ["Fiili Bitiş", "Gerçek Bitiş"],
            "Tamamlanma (%)": ["Tamamlanma (%)", "İlerleme", "İlerleme (%)"], "Ağırlık (%)": ["Ağırlık (%)", "Ağırlık", "Weight"],
            "Sorumlu": ["Sorumlu", "Ekip", "Firma"], "Durum": ["Durum", "Status"], "Kanıt / Saha Notu": ["Kanıt / Saha Notu", "Saha Notu", "Kanıt"]},
    }
    if kind not in mappings:
        raise ValueError(f"Desteklenmeyen içe aktarma türü: {kind}")
    out = pd.DataFrame()
    for target, candidates in mappings[kind].items():
        col = find_column(df, candidates)
        out[target] = df[col] if col else None
    if kind == "plan":
        out["Ana Kategori"] = out["Ana Kategori"].fillna(out["İş Kalemi"].apply(auto_categorize))
        out["Miktar"] = out["Miktar"].apply(safe_num)
        out["Birim Fiyat"] = out["Birim Fiyat"].apply(safe_num)
        for c in ["Başlangıç", "Bitiş"]: out[c] = pd.to_datetime(out[c], errors="coerce", format="mixed", dayfirst=True)
        out = ensure_frame(out, PLAN_COLS)
        return calc_plan_amount(out)
    if kind == "actual":
        out["Gerçekleşen Miktar"] = out["Gerçekleşen Miktar"].apply(safe_num)
        out["Gerçekleşen Birim Fiyat"] = out["Gerçekleşen Birim Fiyat"].apply(safe_num)
        out["Ana Kategori"] = out["Ana Kategori"].fillna(out["İş Kalemi"].apply(auto_categorize))
        out["Tarih"] = pd.to_datetime(out["Tarih"], errors="coerce", format="mixed", dayfirst=True)
        return calc_actual_amount(ensure_frame(out, ACTUAL_COLS))
    out["Süre (gün)"] = out["Süre (gün)"].apply(safe_num)
    out["Tamamlanma (%)"] = out["Tamamlanma (%)"].apply(safe_num).clip(0, 100)
    out["Ağırlık (%)"] = out["Ağırlık (%)"].apply(safe_num)
    out["Bağımlılık Tipi"] = out["Bağımlılık Tipi"].fillna("FS").astype(str).str.upper()
    for c in ["Başlangıç", "Bitiş", "Fiili Başlangıç", "Fiili Bitiş"]: out[c] = pd.to_datetime(out[c], errors="coerce", format="mixed", dayfirst=True)
    return ensure_frame(out, SCHEDULE_COLS)

def read_price_pdf(upload, price_year: int, currency: str) -> tuple[pd.DataFrame, str]:
    data = upload.getvalue()
    if len(data) > 10 * 1024 * 1024:
        raise ValueError("PDF 10 MB sınırını aşıyor. Dosyayı parçalara ayırıp yeniden yükleyin.")
    temp_path = None
    rows = []
    parser_used = ""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(data)
            temp_path = tmp.name
        try:
            from acikpoz import parse_catalog
            parsed = parse_catalog(temp_path)
            for p in getattr(parsed, "pozes", []):
                if getattr(p, "is_priced", False) and getattr(p, "fiyat", None) is not None:
                    rows.append({"Poz No": str(p.poz_no).strip(), "İş Kalemi": str(p.tanim).strip()[:500],
                                 "Birim": str(p.birim).strip(), "Birim Fiyat": safe_num(p.fiyat),
                                 "Para Birimi": currency, "Fiyat Yılı": int(price_year)})
            parser_used = "acikpoz"
        except ImportError:
            try:
                from pypdf import PdfReader
                reader = PdfReader(temp_path)
                text_blob = "\n".join(page.extract_text() or "" for page in reader.pages)
                pattern = re.compile(r"(?m)^\s*([A-ZÇĞİÖŞÜa-zçğıöşü0-9.\-/]{4,})\s+(.{5,}?)\s+(Ad| adet|m|m2|m²|m3|m³|kg|ton|saat|gün|takım|set|%)\s+([\d.,]+)\s*$", re.I)
                for match in pattern.finditer(text_blob):
                    rows.append({"Poz No": match.group(1).strip(), "İş Kalemi": match.group(2).strip(),
                                 "Birim": match.group(3).strip(), "Birim Fiyat": safe_num(match.group(4)),
                                 "Para Birimi": currency, "Fiyat Yılı": int(price_year)})
                parser_used = "pypdf satır eşleştirme (aday kayıt)"
            except ImportError:
                raise RuntimeError("PDF okumak için 'acikpoz' veya 'pypdf' kurulmalı.")
        if not rows:
            raise ValueError("PDF içinden güvenilir fiyatlı poz çıkarılamadı.")
        df = pd.DataFrame(rows)
        df = df[df["Poz No"].astype(str).str.strip().ne("")]
        df = df.drop_duplicates(subset=["Poz No", "Fiyat Yılı"], keep="last")
        return df, parser_used
    finally:
        if temp_path and os.path.exists(temp_path):
            os.unlink(temp_path)

def migrate_v11_legacy_tables() -> list[str]:
    notes: list[str] = []
    try:
        existing_tables = set(sa_inspect(ENGINE).get_table_names())
    except Exception as exc:
        return [f"Eski tablo taraması yapılamadı: {exc}"]
    legacy_names = {"plan", "actual", "schedule", "risk", "cash", "library"}
    if not (existing_tables & legacy_names):
        return notes
    target_project = "V11 - Aktarılan Proje"
    def read_legacy(table: str) -> pd.DataFrame:
        with ENGINE.connect() as con:
            return pd.read_sql(text(f'SELECT * FROM "{table}"'), con)
    migrations = {
        "plan": (PLAN_COLS, {"WBS":"wbs", "Poz No":"poz", "İş Kalemi":"is_kalemi", "Ana Kategori":"ana_kategori", "Alt Kategori":"alt_kategori", "Birim":"birim", "Miktar":"miktar", "Birim Fiyat":"birim_fiyat", "Planlanan Tutar":"planlanan_tutar", "Başlangıç":"baslangic", "Bitiş":"bitis", "Sorumlu":"sorumlu", "Açıklama":"aciklama"}),
        "actual": (ACTUAL_COLS, {"WBS":"wbs", "Poz No":"poz", "İş Kalemi":"is_kalemi", "Ana Kategori":"ana_kategori", "Birim":"birim", "Gerçekleşen Miktar":"gerceklesen_miktar", "Gerçekleşen Birim Fiyat":"gerceklesen_birim_fiyat", "Gerçekleşen Tutar":"gerceklesen_tutar", "Tarih":"tarih", "Tedarikçi":"tedarikci", "Belge No":"belge_no", "Açıklama":"aciklama"}),
        "schedule": (SCHEDULE_COLS, {"WBS":"wbs", "Aktivite":"aktivite", "Süre (gün)":"sure", "Öncül WBS":"oncul", "Bağımlılık Tipi":"bagimlilik", "Başlangıç":"baslangic", "Bitiş":"bitis", "Fiili Başlangıç":"fiili_baslangic", "Fiili Bitiş":"fiili_bitis", "Tamamlanma (%)":"tamamlanma", "Ağırlık (%)":"agirlik", "Sorumlu":"sorumlu", "Durum":"durum", "Kanıt / Saha Notu":"kanit"}),
        "risk": (RISK_COLS, {"Risk ID":"risk_id", "Risk Tanımı":"risk_tanimi", "Kategori":"kategori", "Olasılık (1-5)":"olasilik", "Etki (1-5)":"etki", "Risk Skoru":"risk_skoru", "Önlem / Aksiyon":"onlem", "Sorumlu":"sorumlu", "Durum":"durum", "Tahmini Maliyet Etkisi":"maliyet_etkisi"}),
    }
    for dataset, (cols, mapping) in migrations.items():
        if dataset not in existing_tables:
            continue
        current = load_dataset(target_project, dataset, cols)
        if not current.empty:
            continue
        try:
            raw = read_legacy(dataset)
            if raw.empty:
                continue
            raw = raw.drop(columns=["id"], errors="ignore").copy()
            frame = pd.DataFrame()
            for target, old in mapping.items():
                frame[target] = raw[old] if old in raw.columns else None
            frame = ensure_frame(frame, cols)
            frame["_id"] = [uid() for _ in range(len(frame))]
            save_dataset(target_project, dataset, frame, cols)
            notes.append(f"V11 {dataset} tablosundan {len(frame)} kayıt aktarıldı.")
        except Exception as exc:
            notes.append(f"V11 {dataset} aktarılamadı: {exc}")
    return notes

# ------------------------------ RAPOR / GRAFİK ------------------------------
def excel_export(tables: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, frame in tables.items():
            safe_sheet = re.sub(r"[\[\]:*?/\\]", "_", str(sheet))[:31] or "Rapor"
            data = frame.copy() if isinstance(frame, pd.DataFrame) else pd.DataFrame(frame)
            data.to_excel(writer, sheet_name=safe_sheet, index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            if ws.max_row > 1: ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="14263D")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 32
            for col in ws.columns:
                letter = get_column_letter(col[0].column)
                lengths = [len(str(c.value)) for c in col[:200] if c.value is not None]
                ws.column_dimensions[letter].width = min(max(max(lengths + [10]) + 2, 12), 42)
    output.seek(0)
    return output.getvalue()

def chart_layout(fig, height=390):
    fig.update_layout(height=height, paper_bgcolor="white", plot_bgcolor="white",
                      margin=dict(l=20, r=20, t=45, b=60), font=dict(family="Montserrat", size=11),
                      legend_title_text="")
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#E9EDF2", zeroline=False)
    return fig

def gauge_chart(value, title: str, max_val: float = 2.0, good: float = 1.0):
    val = 0.0 if value is None or pd.isna(value) else float(value)
    fig = go.Figure(go.Indicator(mode="gauge+number", value=val,
        number={"font": {"size": 30, "color": NAVY}, "valueformat": ".2f"},
        title={"text": title, "font": {"size": 13, "color": NAVY}},
        gauge={"axis": {"range": [0, max_val]}, "bar": {"color": TEAL if val >= good else RED},
               "steps": [{"range": [0, good], "color": "#FCE8EA"}, {"range": [good, max_val], "color": "#E5F2EF"}],
               "threshold": {"line": {"color": NAVY, "width": 3}, "value": good}}))
    fig.update_layout(height=260, margin=dict(l=20, r=20, t=50, b=10), paper_bgcolor="white")
    return fig

def calculate_quality_checks(plan, actual, schedule, risk, library) -> pd.DataFrame:
    checks = []
    def add(name, count, detail): checks.append({"Kontrol": name, "Uyarı Sayısı": int(count), "Açıklama": detail})
    add("Bütçede boş WBS", plan["WBS"].fillna("").astype(str).str.strip().eq("").sum(), "WBS bağlantısı EVM ve aktivite takibi için gereklidir.")
    add("Bütçede mükerrer poz/WBS", plan.duplicated(subset=["WBS", "Poz No"], keep=False).sum(), "Aynı WBS + poz çiftini kontrol edin.")
    add("Negatif metraj", (pd.to_numeric(plan["Miktar"], errors="coerce") < 0).sum(), "Miktar negatif olmamalıdır.")
    add("Negatif birim fiyat", (pd.to_numeric(plan["Birim Fiyat"], errors="coerce") < 0).sum(), "Birim fiyat negatif olmamalıdır.")
    pstart, pend = pd.to_datetime(plan["Başlangıç"], errors="coerce", format="mixed", dayfirst=True), pd.to_datetime(plan["Bitiş"], errors="coerce", format="mixed", dayfirst=True)
    add("Bitiş tarihi başlangıçtan önce", ((pend < pstart) & pend.notna() & pstart.notna()).sum(), "Tarih sıralamasını kontrol edin.")
    add("WBS'siz gerçekleşen", actual["WBS"].fillna("").astype(str).str.strip().eq("").sum(), "Fatura kaydını bütçe WBS'ine bağlayın.")
    known = set(plan["WBS"].dropna().astype(str).str.strip())
    add("Bütçede bulunmayan gerçekleşen WBS", (~actual["WBS"].fillna("").astype(str).str.strip().isin(known)).sum(), "Bütçeye eklenmemiş gerçekleşen iş.")
    pr = pd.to_numeric(schedule["Tamamlanma (%)"], errors="coerce")
    add("İlerleme %0–100 dışında", ((pr < 0) | (pr > 100)).sum(), "İlerleme 0–100 aralığında olmalıdır.")
    add("Tekrarlanan aktivite WBS", schedule["WBS"].dropna().astype(str).duplicated().sum(), "CPM için benzersiz WBS.")
    add("Risk puanı eksik", (risk["Risk Skoru"].isna() | risk["Risk Skoru"].eq(0)).sum(), "Olasılık ve etki puanlarını girin.")
    add("Kütüphane poz/fiyat boş", (library["Poz No"].fillna("").astype(str).str.strip().eq("") | pd.to_numeric(library["Birim Fiyat"], errors="coerce").isna()).sum(), "Fiyat kütüphanesi kayıtlarını doğrulayın.")
    return pd.DataFrame(checks)

# ------------------------------ V13: TİCARİ TAHMİN ------------------------------
def tr_norm_text(value: Any) -> str:
    value = str(value).casefold().replace("\u0307", "")
    return value.translate(str.maketrans({"ı":"i", "ş":"s", "ğ":"g", "ü":"u", "ö":"o", "ç":"c"}))

def calculate_employer_hakediş(frame: pd.DataFrame) -> pd.DataFrame:
    out = ensure_frame(frame, EMPLOYER_HAKEDIS_COLS)
    out["Avans Mahsubu"] = out["Avans Mahsubu"].apply(safe_num)
    out["Teminat Kesintisi"] = out["Teminat Kesintisi"].apply(safe_num)
    out["Diğer Kesintiler"] = out["Diğer Kesintiler"].apply(safe_num)
    out["Brüt İmalat Tutarı"] = out["Brüt İmalat Tutarı"].apply(safe_num)
    out["KDV Oranı"] = out["KDV Oranı"].apply(safe_num).clip(0, 1)
    out["Tahsil Edilen"] = out["Tahsil Edilen"].apply(safe_num)
    out["Kesinti Toplamı"] = out["Avans Mahsubu"] + out["Teminat Kesintisi"] + out["Diğer Kesintiler"]
    out["Net Tahakkuk"] = out["Brüt İmalat Tutarı"] - out["Kesinti Toplamı"]
    out["KDV Tutarı"] = out["Net Tahakkuk"] * out["KDV Oranı"]
    out["Hakediş Toplamı"] = out["Net Tahakkuk"] + out["KDV Tutarı"]
    out["Kalan Alacak"] = (out["Hakediş Toplamı"] - out["Tahsil Edilen"]).clip(lower=0)
    return out

def calculate_profit_forecast(plan, actual, schedule, purchase, contracts, changes, evm_values, currency="₺") -> dict:
    base_budget = float(pd.to_numeric(plan.get("Planlanan Tutar", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not plan.empty else 0.0
    actual_cost = float(pd.to_numeric(actual.get("Gerçekleşen Tutar", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not actual.empty else 0.0
    approved_changes = changes.copy() if changes is not None else pd.DataFrame()
    if not approved_changes.empty:
        approved_changes = approved_changes[approved_changes["Durum"].apply(tr_norm_text).isin(["onaylandi", "approved"])]
        if "Para Birimi" in approved_changes.columns:
            change_currency = approved_changes["Para Birimi"].replace("", np.nan).fillna(currency).astype(str).str.strip()
            approved_changes = approved_changes[change_currency == currency]
    approved_revenue_change = pd.to_numeric(approved_changes.get("Gelir Etkisi", pd.Series(dtype=float)), errors="coerce").fillna(0).sum() if not approved_changes.empty else 0.0
    approved_cost_change = pd.to_numeric(approved_changes.get("Maliyet Etkisi", pd.Series(dtype=float)), errors="coerce").fillna(0).sum() if not approved_changes.empty else 0.0
    adjusted_budget = max(base_budget + float(approved_cost_change), 0.0)
    employer_contracts = contracts.copy() if contracts is not None else pd.DataFrame()
    if not employer_contracts.empty:
        type_mask = employer_contracts["Sözleşme Türü"].apply(tr_norm_text).isin(["isveren", "isveren sozlesmesi", "ana sozlesme", "employer"])
        status_mask = employer_contracts["Durum"].apply(tr_norm_text).isin(["onaylandi", "yururlukte", "tamamlandi"])
        contract_currency = employer_contracts["Para Birimi"].replace("", np.nan).fillna(currency).astype(str).str.strip() if "Para Birimi" in employer_contracts.columns else pd.Series(currency, index=employer_contracts.index)
        employer_contracts = employer_contracts[type_mask & status_mask & (contract_currency == currency)]
    contract_revenue = float(pd.to_numeric(employer_contracts.get("Orijinal Bedel", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not employer_contracts.empty else 0.0
    approved_revenue = contract_revenue + float(approved_revenue_change)
    purchase_orders = purchase.copy() if purchase is not None else pd.DataFrame()
    open_commitment = 0.0
    if not purchase_orders.empty:
        valid_status = ~purchase_orders["Durum"].astype(str).str.casefold().isin(["iptal", "cancelled", "talep", "teklif bekleniyor"])
        po = purchase_orders[valid_status].copy()
        if "Para Birimi" in po.columns:
            po_currency = po["Para Birimi"].replace("", np.nan).fillna(currency).astype(str).str.strip()
            po = po[po_currency == currency]
        order_amount = pd.to_numeric(po.get("Sipariş Tutarı", pd.Series(dtype=float)), errors="coerce")
        fallback_amount = pd.to_numeric(po.get("Sipariş Miktarı", pd.Series(dtype=float)), errors="coerce").fillna(0) * pd.to_numeric(po.get("Birim Fiyat", pd.Series(dtype=float)), errors="coerce").fillna(0)
        order_amount = order_amount.fillna(fallback_amount)
        invoiced = pd.to_numeric(po.get("Faturalanan Tutar", pd.Series(dtype=float)), errors="coerce").fillna(0)
        open_commitment = float((order_amount - invoiced).clip(lower=0).sum())
    earned_value = float(evm_values.get("EV", 0) or 0)
    remaining_budget = max(adjusted_budget - earned_value, 0.0)
    etc_bottom_up = max(remaining_budget, open_commitment)
    eac_bottom_up = actual_cost + etc_bottom_up
    cpi = evm_values.get("CPI", np.nan)
    eac_cpi = adjusted_budget / float(cpi) if pd.notna(cpi) and float(cpi) > 0 else np.nan
    candidates = [x for x in [eac_bottom_up, eac_cpi] if pd.notna(x) and math.isfinite(float(x))]
    eac_conservative = max(candidates) if candidates else np.nan
    estimated_profit = approved_revenue - eac_conservative if approved_revenue > 0 and pd.notna(eac_conservative) else np.nan
    margin = estimated_profit / approved_revenue * 100 if approved_revenue > 0 and pd.notna(estimated_profit) else np.nan
    return {
        "Baz Bütçe": base_budget, "Onaylı Maliyet Değişikliği": float(approved_cost_change),
        "Revize Maliyet Bütçesi": adjusted_budget, "Gerçekleşen Maliyet": actual_cost,
        "Kazanılmış Değer (EV)": earned_value, "Kalan Baz Maliyet": remaining_budget,
        "Açık Sipariş Taahhüdü": open_commitment, "Bottom-up ETC": etc_bottom_up,
        "Bottom-up EAC": eac_bottom_up, "CPI Bazlı EAC": float(eac_cpi) if pd.notna(eac_cpi) else np.nan,
        "Muhafazakâr EAC": float(eac_conservative) if pd.notna(eac_conservative) else np.nan,
        "Onaylı İşveren Sözleşmesi": contract_revenue, "Onaylı Değişiklik Geliri": float(approved_revenue_change),
        "Tahmini Sözleşme Geliri": approved_revenue,
        "Tahmini Brüt Kâr": float(estimated_profit) if pd.notna(estimated_profit) else np.nan,
        "Tahmini Kâr Marjı (%)": float(margin) if pd.notna(margin) else np.nan,
    }

# ------------------------------ KULLANICI / ROL ------------------------------
APP_ROLES = ["Yönetici", "Proje Müdürü", "Teknik Ofis", "Şantiye Şefi", "Finans / Muhasebe", "Satın Alma", "Depo / Ambar", "Görüntüleyici"]

def init_auth_db() -> None:
    with ENGINE.begin() as con:
        con.execute(text("""CREATE TABLE IF NOT EXISTS ea_users (
            username TEXT PRIMARY KEY, display_name TEXT NOT NULL, role TEXT NOT NULL,
            salt TEXT NOT NULL, password_hash TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        )"""))

def _password_digest(password: str, salt_hex: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt_hex), 310000).hex()

def list_users() -> pd.DataFrame:
    with ENGINE.connect() as con:
        rows = con.execute(text("SELECT username,display_name,role,active,created_at FROM ea_users ORDER BY username")).fetchall()
    return pd.DataFrame([dict(r._mapping) for r in rows]) if rows else pd.DataFrame(columns=["username","display_name","role","active","created_at"])

def create_user(username: str, display_name: str, role: str, password: str) -> None:
    username = username.strip().lower()
    if not re.fullmatch(r"[a-z0-9_.-]{3,40}", username):
        raise ValueError("Kullanıcı adı 3-40 karakter olmalı.")
    if role not in APP_ROLES:
        raise ValueError("Geçersiz rol.")
    if len(password) < 12:
        raise ValueError("Parola en az 12 karakter olmalı.")
    salt = os.urandom(16).hex()
    digest = _password_digest(password, salt)
    with ENGINE.begin() as con:
        con.execute(text("INSERT INTO ea_users(username,display_name,role,salt,password_hash,active,created_at) VALUES(:u,:d,:r,:s,:h,1,:t)"),
                    {"u":username,"d":display_name.strip() or username,"r":role,"s":salt,"h":digest,"t":now_tr().isoformat(timespec="seconds")})

def reset_user_password(username: str, password: str) -> None:
    if len(password) < 12:
        raise ValueError("Parola en az 12 karakter olmalı.")
    salt = os.urandom(16).hex()
    digest = _password_digest(password, salt)
    with ENGINE.begin() as con:
        result = con.execute(text("UPDATE ea_users SET salt=:s, password_hash=:h WHERE username=:u"), {"s":salt,"h":digest,"u":username})
        if result.rowcount == 0: raise ValueError("Kullanıcı bulunamadı.")

def authenticate_user(username: str, password: str):
    with ENGINE.connect() as con:
        row = con.execute(text("SELECT username,display_name,role,salt,password_hash,active FROM ea_users WHERE username=:u"), {"u":username.strip().lower()}).first()
    if not row or int(row.active) != 1:
        return None
    candidate = _password_digest(password, row.salt)
    if not hmac.compare_digest(candidate, row.password_hash):
        return None
    return {"username":row.username,"display_name":row.display_name,"role":row.role}

def auth_gate() -> None:
    users = list_users()
    if users.empty:
        st.title("🔐 İlk Kurulum — Yönetici Hesabı")
        st.info("İlk yönetici hesabını siz oluşturun. Parola en az 12 karakter olmalıdır.")
        with st.form("bootstrap_admin_form"):
            username = st.text_input("Yönetici kullanıcı adı", value="admin")
            display = st.text_input("Ad Soyad", value="Sistem Yöneticisi")
            p1 = st.text_input("Parola (en az 12 karakter)", type="password")
            p2 = st.text_input("Parolayı tekrar girin", type="password")
            submitted = st.form_submit_button("Yönetici hesabını oluştur", type="primary")
        if submitted:
            if p1 != p2: st.error("Parolalar eşleşmiyor.")
            else:
                try:
                    create_user(username, display, "Yönetici", p1)
                    st.success("Yönetici hesabı oluşturuldu.")
                    st.rerun()
                except Exception as exc: st.error(f"Hesap oluşturulamadı: {exc}")
        st.stop()
    if not st.session_state.get("auth_user"):
        st.title("🔐 EMİR ASEL — Güvenli Giriş")
        with st.form("login_form"):
            username = st.text_input("Kullanıcı adı")
            password = st.text_input("Parola", type="password")
            submitted = st.form_submit_button("Giriş Yap", type="primary")
        if submitted:
            user = authenticate_user(username, password)
            if user:
                st.session_state.auth_user = user
                st.rerun()
            else:
                st.error("Kullanıcı adı veya parola hatalı ya da hesap pasif.")
        st.stop()
    current = st.session_state.get("auth_user", {})
    user_rows = users[users["username"].astype(str) == str(current.get("username", ""))]
    if user_rows.empty or int(user_rows.iloc[0]["active"]) != 1:
        st.session_state.pop("auth_user", None)
        st.error("Kullanıcı hesabı devre dışı bırakılmış.")
        st.stop()
    st.session_state.auth_user["role"] = str(user_rows.iloc[0]["role"])

def role_pages(role: str, all_pages: list[str]) -> list[str]:
    if role == "Yönetici": return all_pages
    common = ["Yönetici Dashboard"]
    access = {
        "Proje Müdürü": common + [x for x in all_pages if x not in ["Kullanıcı Yönetimi"]],
        "Teknik Ofis": common + ["Proje Künyesi / İçe Aktarma", "Birim Fiyat Kütüphanesi", "Bütçe ve Metraj", "Metraj Kontrolü", "İş Programı / CPM", "EVM Performans", "Sözleşme ve Değişiklik Yönetimi", "Günlük Şantiye Raporu", "Onay Merkezi", "Veri Kontrolü", "Excel Raporları"],
        "Şantiye Şefi": common + ["Metraj Kontrolü", "İş Programı / CPM", "İşçilik ve Puantaj", "Depo / Stok", "Günlük Şantiye Raporu", "Risk Yönetimi"],
        "Finans / Muhasebe": common + ["Gerçekleşen Maliyet", "EVM Performans", "Sözleşme ve Değişiklik Yönetimi", "Proje Kârlılığı ve Bitirme Tahmini", "İşveren Hakedişi", "Hakediş Takibi", "Nakit Akışı / S-Curve", "Onay Merkezi", "Excel Raporları"],
        "Satın Alma": common + ["Birim Fiyat Kütüphanesi", "Satın Alma ve Tedarik", "Depo / Stok", "İşçilik ve Puantaj", "Risk Yönetimi"],
        "Depo / Ambar": common + ["Depo / Stok", "Günlük Şantiye Raporu"],
        "Görüntüleyici": common,
    }
    allowed = set(access.get(role, common))
    return [p for p in all_pages if p in allowed]

def add_approval_request(project_key: str, dataset: str, row: pd.Series, title: str, amount: float, currency: str, approver_role: str = "Proje Müdürü") -> str:
    existing = load_dataset(project_key, "approvals", APPROVAL_COLS)
    record_id = str(row.get("_id", ""))
    if not record_id: raise ValueError("Onay talebi için kayıt ID bulunamadı.")
    old_status = str(row.get("Durum", "")).strip().casefold()
    if old_status in ["onaylandı", "onaylandi", "yürürlükte", "yururlukte", "tamamlandı", "tamamlandi"]:
        raise ValueError("Bu kayıt zaten onaylanmış veya yürürlükte.")
    dup = existing[(existing["Veri Seti"].astype(str) == dataset) & (existing["Kayıt ID"].astype(str) == record_id) & (existing["Durum"].astype(str) == "Bekliyor")]
    if not dup.empty: raise ValueError("Bu kayıt için zaten bekleyen bir onay talebi var.")
    user = st.session_state.get("auth_user", {})
    req = {"_id":uid(),"Talep Tarihi":now_tr().isoformat(timespec="seconds"),"Veri Seti":dataset,
           "Modül":dataset,"Kayıt ID":record_id,"Talep Başlığı":title,"Tutar":safe_num(amount),
           "Para Birimi":currency,"Talep Eden":user.get("display_name",user.get("username","")),
           "Onay Rolü":approver_role,"Durum":"Bekliyor","Karar Tarihi":None,"Karar Veren":None,"Karar Notu":""}
    existing = pd.concat([existing, pd.DataFrame([req])], ignore_index=True)
    save_dataset(project_key, "approvals", existing, APPROVAL_COLS)
    st.session_state["data_approvals"] = existing
    if dataset in DATASETS and "Durum" in DATASETS[dataset]:
        target = load_dataset(project_key, dataset, DATASETS[dataset])
        mask = target["_id"].astype(str) == record_id
        if mask.any():
            target.loc[mask, "Durum"] = "Onay Bekliyor"
            save_dataset(project_key, dataset, target, DATASETS[dataset])
            st.session_state[f"data_{dataset}"] = target
    return req["_id"]

def decide_approval(project_key: str, request_id: str, decision: str, note: str = "") -> None:
    approvals = load_dataset(project_key, "approvals", APPROVAL_COLS)
    mask = approvals["_id"].astype(str) == str(request_id)
    if not mask.any(): raise ValueError("Onay talebi bulunamadı.")
    idx = approvals.index[mask][0]
    req = approvals.loc[idx]
    if str(req["Durum"]) != "Bekliyor": raise ValueError("Bu talep daha önce karara bağlanmış.")
    user = st.session_state.get("auth_user", {})
    role = user.get("role", "")
    if role != "Yönetici" and role != str(req["Onay Rolü"]): raise PermissionError("Bu onay talebi için yetkiniz yok.")
    if role != "Yönetici" and str(req["Talep Eden"]) == str(user.get("display_name", user.get("username", ""))):
        raise PermissionError("Talep sahibi kendi talebini onaylayamaz.")
    approvals.loc[idx, "Durum"] = decision
    approvals.loc[idx, "Karar Tarihi"] = now_tr().isoformat(timespec="seconds")
    approvals.loc[idx, "Karar Veren"] = user.get("display_name", user.get("username", ""))
    approvals.loc[idx, "Karar Notu"] = note
    save_dataset(project_key, "approvals", approvals, APPROVAL_COLS)
    st.session_state["data_approvals"] = approvals
    dataset = str(req["Veri Seti"])
    if dataset in DATASETS and "Durum" in DATASETS[dataset]:
        target = load_dataset(project_key, dataset, DATASETS[dataset])
        row_mask = target["_id"].astype(str) == str(req["Kayıt ID"])
        if row_mask.any():
            target.loc[row_mask, "Durum"] = "Onaylandı" if decision == "Onaylandı" else "Reddedildi"
            if "Onay Tarihi" in target.columns: target.loc[row_mask, "Onay Tarihi"] = now_tr().date().isoformat() if decision == "Onaylandı" else None
            if "Onaylayan" in target.columns: target.loc[row_mask, "Onaylayan"] = user.get("display_name", user.get("username", ""))
            save_dataset(project_key, dataset, target, DATASETS[dataset])
            st.session_state[f"data_{dataset}"] = target

# ------------------------------ BAŞLATMA ------------------------------
if DB_ERROR:
    st.error(f"Veritabanı başlatılamadı: {DB_ERROR}")
    st.stop()
else:
    try:
        init_auth_db()
    except Exception as exc:
        st.error(f"Kullanıcı tablosu başlatılamadı: {exc}")
        st.stop()
    auth_gate()
    for note in migrate_v11_legacy_tables():
        st.info(note)

# ------------------------------ PROJE SEÇİMİ ------------------------------
known_projects = load_dataset("__GLOBAL__", "project", PROJECT_COLS)
project_names = sorted(set(known_projects["Proje Adı"].dropna().astype(str).str.strip()) - {""}) if not known_projects.empty else []
default_project = "Yeni Proje"
if "current_project" not in st.session_state:
    st.session_state.current_project = project_names[0] if project_names else default_project

st.markdown(f'<div class="ea-header"><h1>🏗️ {APP_NAME}</h1><div class="sub">ENTEGRE PROJE ERP · SÖZLEŞME · HAKEDİŞ · KÂRLILIK · ŞANTİYE · CPM · EVM · V{APP_VERSION} · {now_tr():%d.%m.%Y %H:%M}</div></div>', unsafe_allow_html=True)

with st.sidebar:
    logo_candidates = [Path(__file__).with_name(x) for x in ["logo.png", "logo.webp", "logo.jpg", "logo.jpeg"]]
    logo_file = next((p for p in logo_candidates if p.exists()), None)
    if logo_file:
        st.image(str(logo_file), width=190)
    st.markdown("## 🏗️ EMİR ASEL")
    st.caption(f"Proje Kontrol Sistemi | V{APP_VERSION}")
    options = list(dict.fromkeys(project_names + [default_project]))
    current_index = options.index(st.session_state.current_project) if st.session_state.current_project in options else len(options)-1
    chosen_project = st.selectbox("Aktif Proje", options, index=current_index)
    if chosen_project != st.session_state.current_project:
        st.session_state.current_project = chosen_project
        st.rerun()
    selected_meta = known_projects[known_projects["Proje Adı"].astype(str) == chosen_project] if not known_projects.empty else pd.DataFrame()
    saved_currency = selected_meta.iloc[0].get("Para Birimi", "₺") if not selected_meta.empty else "₺"
    currency_default = ["₺", "$", "€"].index(saved_currency) if saved_currency in ["₺", "$", "€"] else 0
    currency = st.selectbox("Para Birimi", ["₺", "$", "€"], index=currency_default, key=f"currency_{chosen_project}")
    report_date = st.date_input("Raporlama Tarihi", value=now_tr().date(), key=f"report_date_{chosen_project}")
    st.divider()
    ALL_MENU = ["Yönetici Dashboard", "Proje Künyesi / İçe Aktarma", "Birim Fiyat Kütüphanesi", "Bütçe ve Metraj", "Metraj Kontrolü", "Gerçekleşen Maliyet", "İş Programı / CPM", "EVM Performans", "Satın Alma ve Tedarik", "İşçilik ve Puantaj", "Depo / Stok", "Hakediş Takibi", "Nakit Akışı / S-Curve", "Risk Yönetimi", "Sözleşme ve Değişiklik Yönetimi", "İşveren Hakedişi", "Proje Kârlılığı ve Bitirme Tahmini", "Günlük Şantiye Raporu", "Onay Merkezi", "Veri Kontrolü", "Yedekleme / Denetim", "Excel Raporları", "Kullanıcı Yönetimi"]
    auth_user = st.session_state.get("auth_user", {})
    st.caption(f"Kullanıcı: {auth_user.get('display_name', auth_user.get('username',''))} · Rol: {auth_user.get('role','')}")
    MENU = role_pages(auth_user.get("role", "Görüntüleyici"), ALL_MENU)
    if not MENU: MENU = ["Yönetici Dashboard"]
    page = st.radio("MODÜLLER", MENU)
    if st.button("Oturumu Kapat", width="stretch"):
        st.session_state.pop("auth_user", None)
        st.rerun()
    st.divider()
    st.caption(f"Veritabanı: {'Neon / PostgreSQL' if DB_KIND == 'postgresql' else 'Yerel SQLite'}")

auth_user = st.session_state.get("auth_user", {})
if page not in role_pages(auth_user.get("role", "Görüntüleyici"), ALL_MENU):
    st.error("Bu modüle erişim rolünüz için tanımlı değil.")
    st.stop()

project_key = st.session_state.current_project
if project_key == default_project:
    project_key = "__DRAFT_PROJECT__"

state_key = f"loaded_project::{project_key}"
if state_key not in st.session_state:
    for ds, cols in DATASETS.items():
        if ds == "library":
            st.session_state[f"data_{ds}"] = load_dataset("__GLOBAL__", ds, cols)
        elif ds == "project":
            continue
        else:
            st.session_state[f"data_{ds}"] = load_dataset(project_key, ds, cols)
    st.session_state[state_key] = True

DATA: dict[str, pd.DataFrame] = {ds: st.session_state.get(f"data_{ds}", pd.DataFrame(columns=cols)).copy() for ds, cols in DATASETS.items() if ds != "project"}
for ds, cols in DATASETS.items():
    if ds in DATA:
        DATA[ds] = ensure_frame(DATA[ds], cols)

plan = calc_plan_amount(DATA["plan"])
actual = calc_actual_amount(DATA["actual"])
schedule = DATA["schedule"].copy()
risk = DATA["risk"].copy()
if not risk.empty:
    risk["Risk Skoru"] = pd.to_numeric(risk["Olasılık (1-5)"], errors="coerce").fillna(0) * pd.to_numeric(risk["Etki (1-5)"], errors="coerce").fillna(0)
else:
    risk["Risk Skoru"] = pd.Series(dtype=float)
if not DATA["hakediş"].empty:
    h = DATA["hakediş"].copy()
    h["Net Hakediş"] = pd.to_numeric(h["Brüt Tutar"], errors="coerce").fillna(0) - pd.to_numeric(h["Kesinti"], errors="coerce").fillna(0)
    h["KDV Tutarı"] = h["Net Hakediş"] * pd.to_numeric(h["KDV Oranı"], errors="coerce").fillna(0)
    h["Ödenecek Tutar"] = h["Net Hakediş"] + h["KDV Tutarı"]
    DATA["hakediş"] = h
budget = float(plan["Planlanan Tutar"].sum()) if not plan.empty else 0.0
actual_cost = float(actual["Gerçekleşen Tutar"].sum()) if not actual.empty else 0.0
progress = weighted_progress(schedule)
cost_summary = summarize_costs(plan, actual)
cpm, cpm_warnings = calculate_cpm(schedule)
evm = calculate_evm(plan, actual, schedule, report_date)

# PERFORMANS OPTİMİZASYONU
employer_receivable = 0.0
profit_forecast = {}
if page in ["Yönetici Dashboard", "İşveren Hakedişi", "Proje Kârlılığı ve Bitirme Tahmini", "Excel Raporları"]:
    DATA["employer_hakediş"] = calculate_employer_hakediş(DATA["employer_hakediş"])
    employer_receivable = float(pd.to_numeric(DATA["employer_hakediş"].get("Kalan Alacak", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not DATA["employer_hakediş"].empty else 0.0
if page in ["Yönetici Dashboard", "Proje Kârlılığı ve Bitirme Tahmini", "Excel Raporları"]:
    profit_forecast = calculate_profit_forecast(plan, actual, schedule, DATA["purchase"], DATA["contracts"], DATA["changes"], evm, currency)

def save_from_editor(dataset: str, edited: pd.DataFrame, columns: list[str], recalc: str | None = None):
    frame = ensure_frame(edited, columns)
    if recalc == "plan": frame = calc_plan_amount(frame)
    elif recalc == "actual": frame = calc_actual_amount(frame)
    elif recalc == "risk":
        frame["Risk Skoru"] = pd.to_numeric(frame["Olasılık (1-5)"], errors="coerce").fillna(0) * pd.to_numeric(frame["Etki (1-5)"], errors="coerce").fillna(0)
    elif recalc == "purchase":
        frame["Sipariş Tutarı"] = frame["Sipariş Miktarı"].apply(safe_num) * frame["Birim Fiyat"].apply(safe_num)
    elif recalc == "employer_hakediş":
        frame["Kesinti Toplamı"] = (frame["Avans Mahsubu"].apply(safe_num) + frame["Teminat Kesintisi"].apply(safe_num) + frame["Diğer Kesintiler"].apply(safe_num))
        frame["Net Tahakkuk"] = frame["Brüt İmalat Tutarı"].apply(safe_num) - frame["Kesinti Toplamı"].apply(safe_num)
        frame["KDV Tutarı"] = frame["Net Tahakkuk"] * frame["KDV Oranı"].apply(safe_num)
        frame["Hakediş Toplamı"] = frame["Net Tahakkuk"] + frame["KDV Tutarı"]
        frame["Kalan Alacak"] = (frame["Hakediş Toplamı"] - frame["Tahsil Edilen"].apply(safe_num)).clip(lower=0)
    elif recalc == "labor":
        frame["Toplam İşçilik"] = frame["Çalışan Sayısı"].apply(safe_num) * frame["Gün"].apply(safe_num) * frame["Kişi Günlük Ücret"].apply(safe_num)
    elif recalc == "inventory":
        frame["Tutar"] = frame["Miktar"].apply(safe_num) * frame["Birim Fiyat"].apply(safe_num)
    elif recalc == "hakediş":
        frame["Net Hakediş"] = frame["Brüt Tutar"].apply(safe_num) - frame["Kesinti"].apply(safe_num)
        frame["KDV Tutarı"] = frame["Net Hakediş"] * frame["KDV Oranı"].apply(safe_num)
        frame["Ödenecek Tutar"] = frame["Net Hakediş"] + frame["KDV Tutarı"]
    save_dataset("__GLOBAL__" if dataset == "library" else project_key, dataset, frame, columns)
    st.session_state[f"data_{dataset}"] = frame
    DATA[dataset] = frame
    st.success(f"{len(frame)} kayıt {dataset} kaydedildi.")
    st.rerun()

# ============ SAYFALAR ============
if page == "Yönetici Dashboard":
    st.title(f"Yönetici Dashboard | {st.session_state.current_project}")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Toplam Bütçe", money(budget, currency))
    c2.metric("Gerçekleşen", money(actual_cost, currency), delta=money(actual_cost-budget, currency) if budget else None, delta_color="inverse")
    c3.metric("İlerleme", pct(progress))
    c4.metric("EAC", money(evm["EAC"], currency) if pd.notna(evm["EAC"]) else "—")
    c5,c6,c7,c8 = st.columns(4)
    c5.metric("Kalan Bütçe", money(budget-actual_cost, currency))
    c6.metric("CPI", f'{evm["CPI"]:.2f}' if pd.notna(evm["CPI"]) else "—")
    c7.metric("SPI", f'{evm["SPI"]:.2f}' if pd.notna(evm["SPI"]) else "—")
    c8.metric("Açık Risk", int(risk["Durum"].astype(str).str.casefold().isin(["açık","acik","open"]).sum()) if not risk.empty else 0)
    if auth_user.get("role") in ["Yönetici", "Proje Müdürü", "Finans / Muhasebe"] and profit_forecast:
        st.subheader("Ticari Görünüm")
        f1,f2,f3,f4 = st.columns(4)
        f1.metric("Tahmini Gelir", money(profit_forecast.get("Tahmini Sözleşme Geliri",0), currency) if profit_forecast.get("Tahmini Sözleşme Geliri",0) > 0 else "Sözleşme yok")
        f2.metric("Muhafazakâr EAC", money(profit_forecast.get("Muhafazakâr EAC",np.nan), currency) if pd.notna(profit_forecast.get("Muhafazakâr EAC",np.nan)) else "—")
        f3.metric("Brüt Kâr", money(profit_forecast.get("Tahmini Brüt Kâr",np.nan), currency) if pd.notna(profit_forecast.get("Tahmini Brüt Kâr",np.nan)) else "—")
        f4.metric("İşveren Alacağı", money(employer_receivable, currency))
    left, right = st.columns([1.35,1])
    with left:
        st.subheader("Kategori Bazında Bütçe / Gerçekleşen")
        if not cost_summary.empty:
            chart = cost_summary.melt(id_vars="Ana Kategori", value_vars=["Bütçe","Gerçekleşen"], var_name="Tür", value_name="Tutar")
            fig = px.bar(chart, x="Ana Kategori", y="Tutar", color="Tür", barmode="group", color_discrete_sequence=[NAVY,RED])
            fig.update_layout(xaxis_tickangle=-35, margin=dict(l=10,r=10,t=20,b=130))
            st.plotly_chart(chart_layout(fig, 440), width="stretch")
        else: st.info("Veri yok.")
    with right:
        st.subheader("Bütçe Dağılımı")
        pie = cost_summary[cost_summary["Bütçe"] > 0] if not cost_summary.empty else pd.DataFrame()
        if not pie.empty:
            fig = px.pie(pie, names="Ana Kategori", values="Bütçe", hole=.58, color_discrete_sequence=PALETTE)
            st.plotly_chart(chart_layout(fig, 440), width="stretch")
        else: st.info("Veri yok.")
    st.subheader("Proje Sağlık")
    g1,g2 = st.columns(2)
    with g1: st.plotly_chart(gauge_chart(evm["CPI"], "CPI"), width="stretch")
    with g2: st.plotly_chart(gauge_chart(evm["SPI"], "SPI"), width="stretch")

elif page == "Proje Künyesi / İçe Aktarma":
    st.title("Proje Künyesi ve Veri Aktarımı")
    current_meta = known_projects[known_projects["Proje Adı"].astype(str) == st.session_state.current_project] if not known_projects.empty else pd.DataFrame()
    default_meta = current_meta.iloc[0].to_dict() if not current_meta.empty else {}
    with st.form("project_meta_form"):
        a,b = st.columns(2)
        pname = a.text_input("Proje Adı", value=default_meta.get("Proje Adı", ""))
        owner = b.text_input("İşveren", value=default_meta.get("İşveren", ""))
        c,d = st.columns(2)
        location = c.text_input("Lokasyon", value=default_meta.get("Lokasyon", ""))
        contract_no = d.text_input("Sözleşme No", value=default_meta.get("Sözleşme No", ""))
        e,f = st.columns(2)
        contract_date = e.date_input("Sözleşme Tarihi", value=now_tr().date())
        start_date = f.date_input("Proje Başlangıcı", value=now_tr().date())
        g,h = st.columns(2)
        end_date = g.date_input("Planlanan Bitiş", value=now_tr().date()+timedelta(days=365))
        meta_currency = h.selectbox("Para Birimi", ["₺","$","€"], index=0)
        i,j = st.columns(2)
        vat = i.number_input("KDV Oranı", 0.0, 1.0, 0.20, 0.01)
        reserve = j.number_input("Risk Rezervi (%)", 0.0, 1.0, 0.05, 0.01)
        manager = st.text_input("Proje Yöneticisi", value=default_meta.get("Proje Yöneticisi", ""))
        note = st.text_area("Açıklama", value=default_meta.get("Açıklama", ""))
        submitted = st.form_submit_button("Proje Künyesini Kaydet", type="primary")
        if submitted:
            if not pname.strip(): st.error("Proje adı boş bırakılamaz.")
            else:
                meta = pd.DataFrame([{ "_id": str(default_meta.get("_id") or uid()), "Proje Adı": pname.strip(), "İşveren":owner, "Lokasyon":location,
                    "Sözleşme No":contract_no, "Sözleşme Tarihi":contract_date, "Başlangıç Tarihi":start_date, "Bitiş Tarihi":end_date,
                    "Para Birimi":meta_currency, "KDV Oranı":vat, "Risk Rezervi (%)":reserve, "Proje Yöneticisi":manager, "Açıklama":note }])
                if project_key != pname.strip():
                    migrate_project_records(project_key, pname.strip())
                if not known_projects.empty:
                    excluded_names = {pname.strip(), st.session_state.current_project}
                    new_projects = known_projects[~known_projects["Proje Adı"].astype(str).isin(excluded_names)].copy()
                else:
                    new_projects = pd.DataFrame(columns=PROJECT_COLS)
                new_projects = pd.concat([new_projects, meta], ignore_index=True)
                save_dataset("__GLOBAL__", "project", ensure_frame(new_projects, PROJECT_COLS), PROJECT_COLS)
                st.session_state.current_project = pname.strip()
                st.success("Proje kaydedildi.")
                st.rerun()
    st.divider()
    st.subheader("Excel'den İçe Aktarma")
    uploaded = st.file_uploader("Excel yükle", type=["xlsx","xlsm"], key="project_import_excel")
    if uploaded:
        try:
            sheets = pd.read_excel(uploaded, sheet_name=None)
            sheet_names = list(sheets.keys())
            st.write("Sayfalar:", sheet_names)
            plan_sheet = st.selectbox("Bütçe sayfası", sheet_names)
            actual_sheet = st.selectbox("Gerçekleşen sayfası", ["—"] + sheet_names)
            schedule_sheet = st.selectbox("İş programı sayfası", ["—"] + sheet_names)
            if st.button("İçe aktar", type="primary"):
                p = map_imported_sheet(sheets[plan_sheet], "plan")
                save_dataset(project_key,"plan",p,PLAN_COLS); st.session_state["data_plan"] = p
                if actual_sheet != "—":
                    a = map_imported_sheet(sheets[actual_sheet],"actual")
                    save_dataset(project_key,"actual",a,ACTUAL_COLS); st.session_state["data_actual"] = a
                if schedule_sheet != "—":
                    s = map_imported_sheet(sheets[schedule_sheet],"schedule")
                    save_dataset(project_key,"schedule",s,SCHEDULE_COLS); st.session_state["data_schedule"] = s
                st.success("İçe aktarma tamamlandı.")
                st.rerun()
        except Exception as exc: st.error(f"Hata: {exc}")

elif page == "Birim Fiyat Kütüphanesi":
    st.title("Birim Fiyat Kütüphanesi")
    st.markdown('<div class="ea-note">PDF fiyat listeleri merkezi kütüphaneye kaydedilir. Fiyat yılı ve kaynak korunur.</div>', unsafe_allow_html=True)
    price_year = st.number_input("Fiyat Yılı", min_value=2000, max_value=2100, value=now_tr().year, step=1)
    pdf_currency = st.selectbox("PDF Para Birimi", ["₺","$","€"], key="pdf_currency")
    pdf_upload = st.file_uploader("PDF yükle (maksimum 10 MB)", type=["pdf"], key="price_pdf")
    if pdf_upload and st.button("PDF'i Oku", type="primary"):
        try:
            parsed, parser_name = read_price_pdf(pdf_upload, int(price_year), pdf_currency)
            digest = hashlib.sha256(pdf_upload.getvalue()).hexdigest()[:16]
            saved_name = f"{int(price_year)}_{digest}_{Path(pdf_upload.name).name}"
            (PDF_DIR / saved_name).write_bytes(pdf_upload.getvalue())
            parsed["Kaynak"] = f"PDF: {pdf_upload.name} ({parser_name})"
            parsed["Kaynak Dosya"] = saved_name
            parsed["Yayın / Güncelleme Tarihi"] = now_tr().date().isoformat()
            parsed["Aktif"] = True
            parsed = ensure_frame(parsed, LIB_COLS)
            existing = DATA["library"].copy()
            keys = set(zip(existing["Poz No"].astype(str), existing["Fiyat Yılı"].astype(str), existing["Kaynak Dosya"].astype(str))) if not existing.empty else set()
            parsed = parsed[[ (str(r["Poz No"]), str(r["Fiyat Yılı"]), str(r["Kaynak Dosya"])) not in keys for _,r in parsed.iterrows() ]]
            combined = pd.concat([existing, parsed], ignore_index=True)
            save_dataset("__GLOBAL__","library",combined,LIB_COLS)
            st.session_state["data_library"] = ensure_frame(combined, LIB_COLS)
            st.success(f"{len(parsed)} poz eklendi. Parser: {parser_name}")
            st.rerun()
        except Exception as exc: st.error(f"Hata: {exc}")
    lib = DATA["library"].copy()
    search = st.text_input("Ara (Poz / İş Kalemi / Kaynak)", key="library_search")
    if search:
        mask = lib.astype(str).apply(lambda c:c.str.contains(search,case=False,na=False)).any(axis=1)
        lib = lib[mask]
    st.dataframe(lib.drop(columns=["_id"], errors="ignore"), width="stretch", hide_index=True, height=430)

elif page == "Bütçe ve Metraj":
    st.title("Bütçe ve Metraj")
    edited = st.data_editor(plan, num_rows="dynamic", width="stretch", hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Ana Kategori":st.column_config.SelectboxColumn(options=COST_CATEGORIES),
            "Metraj Durumu":st.column_config.SelectboxColumn(options=["Taslak","Kontrol Bekliyor","Onaylandı","Revize","Tahmini"]),
            "Miktar":st.column_config.NumberColumn(min_value=0,format="%.3f"),
            "Birim Fiyat":st.column_config.NumberColumn(min_value=0,format="%.4f"),
            "Planlanan Tutar":st.column_config.NumberColumn(disabled=True,format="%.2f"),
            "Başlangıç":st.column_config.DateColumn(),"Bitiş":st.column_config.DateColumn()}, key="plan_editor")
    if st.button("Kaydet", type="primary"):
        save_from_editor("plan",edited,PLAN_COLS,"plan")
    if not cost_summary.empty:
        st.subheader("Kategori Özeti"); st.dataframe(cost_summary,width="stretch",hide_index=True)
    st.metric("Toplam Bütçe",money(budget,currency))

elif page == "Metraj Kontrolü":
    st.title("Metraj Kontrol")
    c1,c2,c3 = st.columns(3)
    pending = plan[plan["Metraj Durumu"].fillna("Taslak").isin(["Taslak","Kontrol Bekliyor","Revize","Tahmini"])]
    approved = plan[plan["Metraj Durumu"].fillna("").eq("Onaylandı")]
    c1.metric("Toplam",len(plan)); c2.metric("Bekleyen",len(pending)); c3.metric("Onaylı",len(approved))
    if not plan.empty:
        st.dataframe(plan.drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page == "Gerçekleşen Maliyet":
    st.title("Gerçekleşen Maliyet")
    edited = st.data_editor(actual,num_rows="dynamic",width="stretch",hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Ana Kategori":st.column_config.SelectboxColumn(options=COST_CATEGORIES),
            "Kayıt Türü":st.column_config.SelectboxColumn(options=["Fatura","Taşeron Hakedişi","Malzeme","Hizmet","Diğer"]),
            "Gerçekleşen Miktar":st.column_config.NumberColumn(min_value=0,format="%.3f"),
            "Gerçekleşen Birim Fiyat":st.column_config.NumberColumn(min_value=0,format="%.4f"),
            "Gerçekleşen Tutar":st.column_config.NumberColumn(disabled=True,format="%.2f"),"Tarih":st.column_config.DateColumn()}, key="actual_editor")
    if st.button("Kaydet",type="primary"): save_from_editor("actual",edited,ACTUAL_COLS,"actual")
    st.dataframe(cost_summary,width="stretch",hide_index=True)

elif page == "İş Programı / CPM":
    st.title("İş Programı / CPM")
    st.markdown('<div class="ea-note">Öncül: A01 (FS), A01:SS:2 gibi. FS/SS/FF/SF desteklenir.</div>',unsafe_allow_html=True)
    edited = st.data_editor(schedule,num_rows="dynamic",width="stretch",hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Süre (gün)":st.column_config.NumberColumn(min_value=0,step=1),
            "Bağımlılık Tipi":st.column_config.SelectboxColumn(options=["FS","SS","FF","SF"]),
            "Tamamlanma (%)":st.column_config.NumberColumn(min_value=0,max_value=100,format="%.1f"),
            "Ağırlık (%)":st.column_config.NumberColumn(min_value=0,format="%.2f"),
            "Başlangıç":st.column_config.DateColumn(),"Bitiş":st.column_config.DateColumn(),
            "Fiili Başlangıç":st.column_config.DateColumn(),"Fiili Bitiş":st.column_config.DateColumn()},key="schedule_editor")
    if st.button("Kaydet",type="primary"):
        frame=ensure_frame(edited,SCHEDULE_COLS)
        frame["Süre (gün)"]=pd.to_numeric(frame["Süre (gün)"],errors="coerce").fillna(0).clip(lower=0)
        frame["Tamamlanma (%)"]=pd.to_numeric(frame["Tamamlanma (%)"],errors="coerce").fillna(0).clip(0,100)
        frame["Ağırlık (%)"]=pd.to_numeric(frame["Ağırlık (%)"],errors="coerce").fillna(0).clip(lower=0)
        save_from_editor("schedule",frame,SCHEDULE_COLS)
    for w in cpm_warnings: st.warning(w)
    if not cpm.empty:
        st.metric("Proje Süresi",f'{cpm["Erken Bitiş"].max():.0f} gün')
        st.dataframe(cpm.drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)
        fig=px.bar(cpm,x="Aktivite",y="Süre (gün)",color="Kritik mi?",color_discrete_map={"EVET":RED,"Hayır":TEAL})
        st.plotly_chart(chart_layout(fig,400),width="stretch")
        dated=schedule.dropna(subset=["Başlangıç","Bitiş"]).copy()
        if not dated.empty:
            fig=px.timeline(dated,x_start="Başlangıç",x_end="Bitiş",y="Aktivite",color="Tamamlanma (%)",color_continuous_scale="Blues")
            fig.update_yaxes(autorange="reversed")
            st.plotly_chart(chart_layout(fig,450),width="stretch")

elif page == "EVM Performans":
    st.title("Earned Value Management")
    st.caption(f"Rapor tarihi: {report_date:%d.%m.%Y}")
    cols=st.columns(5)
    for col,label in zip(cols,["BAC","PV","EV","AC","EAC"]):
        col.metric(label,money(evm[label],currency) if pd.notna(evm[label]) else "—")
    c1,c2,c3,c4=st.columns(4)
    c1.metric("CPI",f'{evm["CPI"]:.3f}' if pd.notna(evm["CPI"]) else "—")
    c2.metric("SPI",f'{evm["SPI"]:.3f}' if pd.notna(evm["SPI"]) else "—")
    c3.metric("CV",money(evm["CV"],currency))
    c4.metric("SV",money(evm["SV"],currency))
    st.dataframe(pd.DataFrame([{"Gösterge":k,"Değer":v} for k,v in evm.items()]),width="stretch",hide_index=True)

elif page == "Satın Alma ve Tedarik":
    st.title("Satın Alma ve Tedarik")
    edited=st.data_editor(DATA["purchase"],num_rows="dynamic",width="stretch",hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Sipariş Tarihi":st.column_config.DateColumn(),"Termin Tarihi":st.column_config.DateColumn(),"Fatura Tarihi":st.column_config.DateColumn(),
            "Sipariş Miktarı":st.column_config.NumberColumn(min_value=0,format="%.3f"),"Teslim Miktarı":st.column_config.NumberColumn(min_value=0,format="%.3f"),
            "Birim Fiyat":st.column_config.NumberColumn(min_value=0,format="%.4f"),"Para Birimi":st.column_config.SelectboxColumn(options=["₺","$","€"]),"Sipariş Tutarı":st.column_config.NumberColumn(disabled=True,format="%.2f"),
            "Durum":st.column_config.SelectboxColumn(options=["Talep","Teklif Bekleniyor","Sipariş Verildi","Kısmi Teslimat","Teslim Edildi","İptal"] )},key="purchase_editor")
    if st.button("Kaydet",type="primary"): save_from_editor("purchase",edited,PURCHASE_COLS,"purchase")
    pur=DATA["purchase"].copy()
    if not pur.empty:
        st.dataframe(pur.drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)
        st.divider(); st.subheader("Faturayı Gerçekleşen Maliyete Aktar")
        invoice_options={f'{r.get("Sipariş No", "")} · {r.get("Tedarikçi", "")} · {r.get("Fatura No", "No girilmedi")}':str(r["_id"]) for _,r in pur.iterrows()}
        chosen_invoice=st.selectbox("Fatura",list(invoice_options),key="purchase_invoice_post_choice")
        if st.button("Aktar",key="post_purchase_invoice"):
            purchase_row=pur[pur["_id"].astype(str)==invoice_options[chosen_invoice]].iloc[0]
            invoice_no=str(purchase_row.get("Fatura No", "") or "").strip()
            invoice_amount=safe_num(purchase_row.get("Faturalanan Tutar"))
            if not invoice_no or invoice_no.casefold() in ["nan","none"]:
                st.error("Fatura No doldurun.")
            elif invoice_amount<=0:
                st.error("Tutar > 0 olmalı.")
            elif not actual.empty and actual["Belge No"].astype(str).str.strip().eq(invoice_no).any():
                st.warning("Bu fatura zaten aktarılmış.")
            else:
                new_actual={"_id":uid(),"WBS":purchase_row.get("WBS"),"Poz No":purchase_row.get("Poz No"),
                    "İş Kalemi":purchase_row.get("Malzeme / Hizmet"),"Ana Kategori":auto_categorize(str(purchase_row.get("Malzeme / Hizmet",""))),
                    "Birim":purchase_row.get("Birim"),"Gerçekleşen Miktar":1,"Gerçekleşen Birim Fiyat":invoice_amount,
                    "Gerçekleşen Tutar":invoice_amount,"Tarih":purchase_row.get("Fatura Tarihi"),"Tedarikçi":purchase_row.get("Tedarikçi"),
                    "Belge No":invoice_no,"Kayıt Türü":"Fatura","Açıklama":f'{purchase_row.get("Sipariş No", "")}'}
                merged_actual=pd.concat([actual,pd.DataFrame([new_actual])],ignore_index=True)
                save_dataset(project_key,"actual",merged_actual,ACTUAL_COLS)
                st.success("Aktarıldı."); st.rerun()

elif page == "İşçilik ve Puantaj":
    st.title("İşçilik ve Puantaj")
    edited=st.data_editor(DATA["labor"],num_rows="dynamic",width="stretch",hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Tarih":st.column_config.DateColumn(),"Çalışan Sayısı":st.column_config.NumberColumn(min_value=0),"Gün":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Kişi Günlük Ücret":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Toplam İşçilik":st.column_config.NumberColumn(disabled=True,format="%.2f")},key="labor_editor")
    if st.button("Kaydet",type="primary"): save_from_editor("labor",edited,LABOR_COLS,"labor")
    labor=DATA["labor"]
    if not labor.empty:
        total=pd.to_numeric(labor["Toplam İşçilik"],errors="coerce").fillna(0).sum()
        st.metric("Toplam İşçilik",money(total,currency))

elif page == "Depo / Stok":
    st.title("Depo ve Stok")
    edited=st.data_editor(DATA["inventory"],num_rows="dynamic",width="stretch",hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Tarih":st.column_config.DateColumn(),"Hareket Türü":st.column_config.SelectboxColumn(options=["Giriş","Çıkış","Sayım Artısı","Sayım Eksisi","Düzeltme"]),"Miktar":st.column_config.NumberColumn(min_value=0,format="%.3f"),"Birim Fiyat":st.column_config.NumberColumn(min_value=0,format="%.4f"),"Tutar":st.column_config.NumberColumn(disabled=True,format="%.2f")},key="inventory_editor")
    if st.button("Kaydet",type="primary"): save_from_editor("inventory",edited,INVENTORY_COLS,"inventory")
    inv=DATA["inventory"].copy()
    if not inv.empty:
        inv["Miktar"] = pd.to_numeric(inv["Miktar"],errors="coerce").fillna(0)
        movement = inv["Hareket Türü"].astype(str)
        sign = np.where(movement.isin(["Çıkış", "Sayım Eksisi"]), -1, 1)
        inv["Net Miktar"] = sign * inv["Miktar"]
        inv["Net Değer"] = sign * pd.to_numeric(inv["Tutar"],errors="coerce").fillna(0)
        balances=inv.groupby(["Depo","Malzeme Kodu","Malzeme","Birim"],dropna=False).agg(Stok=("Net Miktar","sum"),Stok_Değeri=("Net Değer","sum")).reset_index()
        st.dataframe(balances,width="stretch",hide_index=True)
        st.dataframe(inv.drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page == "Hakediş Takibi":
    st.title("Taşeron Hakediş")
    edited=st.data_editor(DATA["hakediş"],num_rows="dynamic",width="stretch",hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Tarih":st.column_config.DateColumn(),"Brüt Tutar":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Kesinti":st.column_config.NumberColumn(min_value=0,format="%.2f"),"KDV Oranı":st.column_config.NumberColumn(min_value=0,max_value=1,format="%.2f"),"Net Hakediş":st.column_config.NumberColumn(disabled=True,format="%.2f"),"KDV Tutarı":st.column_config.NumberColumn(disabled=True,format="%.2f"),"Ödenecek Tutar":st.column_config.NumberColumn(disabled=True,format="%.2f"),"Ödenen":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Durum":st.column_config.SelectboxColumn(options=["Hazırlanıyor","Kontrolde","Onaylandı","Kısmi Ödendi","Ödendi","İtirazlı"] )},key="hak_editor")
    if st.button("Kaydet",type="primary"): save_from_editor("hakediş",edited,HAKEDIS_COLS,"hakediş")
    h=DATA["hakediş"].copy()
    if not h.empty:
        st.dataframe(h.drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page == "Nakit Akışı / S-Curve":
    st.title("Nakit Akışı / S-Curve")
    edited=st.data_editor(DATA["cash"],num_rows="dynamic",width="stretch",hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Tarih":st.column_config.DateColumn(),"Vade Tarihi":st.column_config.DateColumn(),"Tür":st.column_config.SelectboxColumn(options=["Tahsilat","Ödeme"]),"Tutar":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Durum":st.column_config.SelectboxColumn(options=["Planlandı","Gerçekleşti","Gecikti"] )},key="cash_editor")
    if st.button("Kaydet",type="primary"):
        frame=ensure_frame(edited,CASH_COLS); frame["Tutar"]=pd.to_numeric(frame["Tutar"],errors="coerce").fillna(0); save_from_editor("cash",frame,CASH_COLS)
    cash=DATA["cash"].copy()
    if not cash.empty:
        cash["Tutar"]=pd.to_numeric(cash["Tutar"],errors="coerce").fillna(0)
        realized=cash[cash["Durum"].astype(str)=="Gerçekleşti"]
        inflow=realized.loc[realized["Tür"]=="Tahsilat","Tutar"].sum(); outflow=realized.loc[realized["Tür"]=="Ödeme","Tutar"].sum()
        a,b,c=st.columns(3); a.metric("Tahsilat",money(inflow,currency)); b.metric("Ödeme",money(outflow,currency)); c.metric("Net",money(inflow-outflow,currency))

elif page == "Risk Yönetimi":
    st.title("Risk Yönetimi")
    edited=st.data_editor(risk.drop(columns=["Risk Skoru"],errors="ignore"),num_rows="dynamic",width="stretch",hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Olasılık (1-5)":st.column_config.NumberColumn(min_value=1,max_value=5,step=1),"Etki (1-5)":st.column_config.NumberColumn(min_value=1,max_value=5,step=1),"Tahmini Maliyet Etkisi":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Durum":st.column_config.SelectboxColumn(options=["Açık","İzleniyor","Kapalı"]),"Hedef Tarih":st.column_config.DateColumn()},key="risk_editor")
    if st.button("Kaydet",type="primary"): save_from_editor("risk",edited,RISK_COLS,"risk")
    if not risk.empty:
        fig=px.scatter(risk,x="Olasılık (1-5)",y="Etki (1-5)",size="Tahmini Maliyet Etkisi",color="Risk Skoru",hover_name="Risk Tanımı",text="Risk ID",range_x=[0.5,5.5],range_y=[0.5,5.5],color_continuous_scale="RdYlGn_r")
        st.plotly_chart(chart_layout(fig,480),width="stretch")

elif page == "Sözleşme ve Değişiklik Yönetimi":
    st.title("Sözleşme ve Değişiklik Yönetimi")
    st.subheader("Sözleşmeler")
    contract_edit = st.data_editor(DATA["contracts"], num_rows="dynamic", width="stretch", hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Sözleşme Türü":st.column_config.SelectboxColumn(options=["İşveren", "Alt Yüklenici", "Tedarik", "Danışmanlık", "Diğer"]),
            "Sözleşme Tarihi":st.column_config.DateColumn(),"Başlangıç Tarihi":st.column_config.DateColumn(),"Bitiş Tarihi":st.column_config.DateColumn(),
            "Para Birimi":st.column_config.SelectboxColumn(options=["₺","$","€"]),"Orijinal Bedel":st.column_config.NumberColumn(min_value=0,format="%.2f")}, key="contracts_editor")
    if st.button("Sözleşmeleri Kaydet",type="primary",key="save_contracts"):
        save_from_editor("contracts",contract_edit,CONTRACT_COLS)
    st.subheader("Değişiklikler")
    change_edit = st.data_editor(DATA["changes"], num_rows="dynamic", width="stretch", hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Talep Tarihi":st.column_config.DateColumn(),
            "Gelir Etkisi":st.column_config.NumberColumn(format="%.2f"),"Maliyet Etkisi":st.column_config.NumberColumn(format="%.2f"),
            "Süre Etkisi (gün)":st.column_config.NumberColumn(step=1),"Para Birimi":st.column_config.SelectboxColumn(options=["₺","$","€"])}, key="changes_editor")
    if st.button("Değişiklikleri Kaydet",type="primary",key="save_changes"):
        save_from_editor("changes",change_edit,CHANGE_COLS)

elif page == "İşveren Hakedişi":
    st.title("İşveren Hakedişi, Faturalama, Tahsilat")
    employer_edit=st.data_editor(DATA["employer_hakediş"],num_rows="dynamic",width="stretch",hide_index=True,
        column_config={"_id":st.column_config.Column(hidden=True),"Tarih":st.column_config.DateColumn(),"Vade Tarihi":st.column_config.DateColumn(),
            "Brüt İmalat Tutarı":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Avans Mahsubu":st.column_config.NumberColumn(min_value=0,format="%.2f"),
            "Teminat Kesintisi":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Diğer Kesintiler":st.column_config.NumberColumn(min_value=0,format="%.2f"),
            "KDV Oranı":st.column_config.NumberColumn(min_value=0,max_value=1,format="%.2f"),"Tahsil Edilen":st.column_config.NumberColumn(min_value=0,format="%.2f"),
            "Kesinti Toplamı":st.column_config.NumberColumn(disabled=True,format="%.2f"),"KDV Tutarı":st.column_config.NumberColumn(disabled=True,format="%.2f"),
            "Net Tahakkuk":st.column_config.NumberColumn(disabled=True,format="%.2f"),"Hakediş Toplamı":st.column_config.NumberColumn(disabled=True,format="%.2f"),
            "Kalan Alacak":st.column_config.NumberColumn(disabled=True,format="%.2f")},key="employer_hakediş_editor")
    if st.button("Kaydet",type="primary"):
        save_from_editor("employer_hakediş",employer_edit,EMPLOYER_HAKEDIS_COLS,"employer_hakediş")
    e_h=calculate_employer_hakediş(DATA["employer_hakediş"])
    if not e_h.empty:
        k1,k2,k3,k4=st.columns(4)
        k1.metric("Brüt",money(e_h["Brüt İmalat Tutarı"].sum(),currency))
        k2.metric("Hakediş",money(e_h["Hakediş Toplamı"].sum(),currency))
        k3.metric("Tahsilat",money(e_h["Tahsil Edilen"].sum(),currency))
        k4.metric("Açık Alacak",money(e_h["Kalan Alacak"].sum(),currency))

elif page == "Proje Kârlılığı ve Bitirme Tahmini":
    st.title("Proje Kârlılığı ve Bitirme Maliyeti Tahmini")
    f=profit_forecast
    a,b,c,d=st.columns(4)
    a.metric("Tahmini Gelir",money(f.get("Tahmini Sözleşme Geliri",0),currency) if f.get("Tahmini Sözleşme Geliri",0)>0 else "Sözleşme yok")
    b.metric("Revize Bütçe",money(f.get("Revize Maliyet Bütçesi",0),currency))
    c.metric("Gerçekleşen",money(f.get("Gerçekleşen Maliyet",0),currency))
    d.metric("Açık Taahhüt",money(f.get("Açık Sipariş Taahhüdü",0),currency))
    st.subheader("EAC Senaryoları")
    s1,s2,s3=st.columns(3)
    s1.metric("Bottom-up EAC",money(f.get("Bottom-up EAC",0),currency))
    s2.metric("CPI Bazlı EAC",money(f.get("CPI Bazlı EAC",np.nan),currency) if pd.notna(f.get("CPI Bazlı EAC",np.nan)) else "CPI eksik")
    s3.metric("Muhafazakâr EAC",money(f.get("Muhafazakâr EAC",np.nan),currency) if pd.notna(f.get("Muhafazakâr EAC",np.nan)) else "—")
    st.dataframe(pd.DataFrame([{"Gösterge":k,"Değer":v} for k,v in f.items()]),width="stretch",hide_index=True)

elif page == "Günlük Şantiye Raporu":
    st.title("Günlük Şantiye Raporu")
    with st.form("daily_site_report_form",clear_on_submit=True):
        r1,r2,r3=st.columns(3)
        report_day=r1.date_input("Rapor tarihi",value=now_tr().date())
        weather=r2.selectbox("Hava",["Açık","Parçalı bulutlu","Bulutlu","Yağmurlu","Fırtınalı","Aşırı sıcak","Diğer"])
        temp=r3.number_input("Sıcaklık (°C)",value=20.0,step=1.0)
        r4,r5=st.columns(2)
        site_manager=r4.text_input("Şantiye sorumlusu",value=auth_user.get("display_name",""))
        workers=r5.number_input("Personel",min_value=0,step=1,value=0)
        teams=st.text_area("Ekipler",height=70)
        equipment=st.text_area("Makineler",height=70)
        activities=st.text_area("İmalat özeti",height=110)
        q1,q2,q3,q4=st.columns(4)
        concrete=q1.number_input("Beton (m³)",min_value=0.0,value=0.0,step=1.0)
        rebar=q2.number_input("Donatı (ton)",min_value=0.0,value=0.0,step=0.1)
        formwork=q3.number_input("Kalıp (m²)",min_value=0.0,value=0.0,step=1.0)
        excavation=q4.number_input("Kazı (m³)",min_value=0.0,value=0.0,step=1.0)
        deliveries=st.text_area("Gelen malzemeler",height=70)
        delays=st.text_area("Gecikmeler",height=70)
        safety=st.text_area("İSG",height=70)
        tomorrow=st.text_area("Yarınki plan",height=70)
        notes=st.text_area("Açıklama",height=70)
        files=st.file_uploader("Foto/PDF ekle",type=["jpg","jpeg","png","pdf"],accept_multiple_files=True)
        submitted=st.form_submit_button("Kaydet",type="primary")
    if submitted:
        # BELLEK KORUMASI: Toplam dosya boyutu 50 MB
        if files:
            total_size = sum(f.size for f in files)
            if total_size > 50 * 1024 * 1024:
                st.error(f"Toplam dosya boyutu çok büyük: {total_size/1024/1024:.1f} MB. 50 MB'tan az yükleyin.")
                st.stop()
        report_id=uid()
        safe_project=re.sub(r"[^A-Za-z0-9_-]+","_",st.session_state.current_project)[:60]
        report_dir=Path(__file__).with_name("data")/"daily_reports"/safe_project
        report_dir.mkdir(parents=True,exist_ok=True)
        saved_files=[]
        for i,uploaded_file in enumerate(files or []):
            if uploaded_file.size > 10 * 1024 * 1024:
                st.warning(f"{uploaded_file.name} 10 MB sınırını aşıyor, atlandı.")
                continue
            suffix=Path(uploaded_file.name).suffix.lower()
            filename=f"{report_day:%Y%m%d}_{report_id[:10]}_{i+1}{suffix}"
            target=report_dir/filename
            target.write_bytes(uploaded_file.getvalue())
            saved_files.append(str(target.relative_to(Path(__file__).parent)))
        rec={"_id":report_id,"Rapor Tarihi":report_day.isoformat(),"Hava Durumu":weather,"Sıcaklık (°C)":temp,
             "Şantiye Sorumlusu":site_manager,"Toplam Personel":workers,"Alt Yükleniciler / Ekipler":teams,"Çalışan Makineler":equipment,
             "Günlük İmalat Özeti":activities,"Beton (m³)":concrete,"Donatı (ton)":rebar,"Kalıp (m²)":formwork,"Kazı / Dolgu (m³)":excavation,
             "Gelen Malzemeler":deliveries,"Gecikme / Engel":delays,"İSG / Olay Bilgisi":safety,"Yarınki Plan":tomorrow,"Açıklama":notes,
             "Ek Dosyalar":json.dumps(saved_files,ensure_ascii=False),"Raporlayan":auth_user.get("display_name","")}
        updated=pd.concat([DATA["daily_reports"],pd.DataFrame([rec])],ignore_index=True)
        save_dataset(project_key,"daily_reports",updated,DAILY_REPORT_COLS)
        st.success("Rapor kaydedildi."); st.rerun()
    reports=DATA["daily_reports"].copy()
    if not reports.empty:
        st.dataframe(reports.drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page == "Onay Merkezi":
    st.title("Onay Merkezi")
    user=auth_user
    approvals=DATA["approvals"].copy()
    if approvals.empty:
        st.info("Onay talebi yok.")
    else:
        pending=approvals[approvals["Durum"].astype(str)=="Bekliyor"].copy()
        st.metric("Bekleyen",len(pending))
        st.dataframe(pending.drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)
        if not pending.empty and user.get("role") in ["Yönetici","Proje Müdürü"]:
            allowed=pending if user.get("role")=="Yönetici" else pending[pending["Onay Rolü"].astype(str)==user.get("role")]
            if not allowed.empty:
                opts={f'{r.get("Talep Başlığı","")} · {money(r.get("Tutar",0))}':str(r["_id"]) for _,r in allowed.iterrows()}
                chosen=st.selectbox("Talep",list(opts),key="approval_decision_choice")
                note=st.text_area("Karar notu",key="approval_decision_note")
                c1,c2=st.columns(2)
                if c1.button("Onayla",type="primary"):
                    try: decide_approval(project_key,opts[chosen],"Onaylandı",note); st.rerun()
                    except Exception as exc: st.error(str(exc))
                if c2.button("Reddet"):
                    try: decide_approval(project_key,opts[chosen],"Reddedildi",note); st.rerun()
                    except Exception as exc: st.error(str(exc))
        st.subheader("Tüm Geçmiş")
        st.dataframe(approvals.drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page == "Kullanıcı Yönetimi":
    st.title("Kullanıcı Yönetimi")
    if auth_user.get("role") != "Yönetici":
        st.error("Sadece Yönetici erişebilir.")
        st.stop()
    with st.form("create_app_user_form"):
        u1,u2,u3=st.columns(3)
        username=u1.text_input("Kullanıcı adı")
        display=u2.text_input("Ad Soyad")
        role=u3.selectbox("Rol",APP_ROLES,index=1)
        pw1=st.text_input("Parola (min 12)",type="password")
        pw2=st.text_input("Tekrar",type="password")
        if st.form_submit_button("Oluştur",type="primary"):
            if pw1!=pw2: st.error("Parolalar eşleşmiyor.")
            else:
                try: create_user(username,display,role,pw1); st.success("Oluşturuldu."); st.rerun()
                except Exception as exc: st.error(f"Hata: {exc}")
    users=list_users()
    st.dataframe(users,width="stretch",hide_index=True)
    if not users.empty:
        selected_user=st.selectbox("Düzenle",users["username"].tolist(),key="edit_user_choice")
        target=users[users["username"]==selected_user].iloc[0]
        new_role=st.selectbox("Rol",APP_ROLES,index=APP_ROLES.index(target["role"]) if target["role"] in APP_ROLES else 0,key="edit_user_role")
        active=st.checkbox("Aktif",value=bool(int(target["active"])),key="edit_user_active")
        if st.button("Kaydet",key="save_user_settings"):
            if selected_user==auth_user.get("username") and not active:
                st.error("Kendi hesabınızı devre dışı bırakamazsınız.")
            else:
                with ENGINE.begin() as con:
                    con.execute(text("UPDATE ea_users SET role=:r, active=:a WHERE username=:u"),{"r":new_role,"a":1 if active else 0,"u":selected_user})
                st.success("Kaydedildi."); st.rerun()

elif page == "Veri Kontrolü":
    st.title("Veri Kontrolü")
    checks=calculate_quality_checks(plan,actual,schedule,risk,DATA["library"])
    st.dataframe(checks,width="stretch",hide_index=True)
    total=int(checks["Uyarı Sayısı"].sum()) if not checks.empty else 0
    if total==0: st.success("Uyarı yok.")
    else: st.warning(f"{total} uyarı.")

elif page == "Yedekleme / Denetim":
    st.title("Yedekleme / Denetim")
    snapshot_sets={ds:DATA[ds] for ds in DATA if ds!="library"}
    if st.button("Snapshot oluştur",type="primary"):
        snap_id=make_snapshot(project_key,snapshot_sets)
        st.success(f"Snapshot: {snap_id}")
    export_json={"app":APP_NAME,"version":APP_VERSION,"project":st.session_state.current_project,"created_at":now_tr().isoformat(),"data":{k:v.to_dict(orient="records") for k,v in snapshot_sets.items()},"library":DATA["library"].to_dict(orient="records")}
    st.download_button("JSON Yedek İndir",json.dumps(json_safe_value(export_json),ensure_ascii=False,allow_nan=False,indent=2).encode("utf-8"),file_name=f"Emir_Asel_backup_{now_tr():%Y%m%d_%H%M}.json",mime="application/json")
    st.subheader("Denetim Kayıtları")
    logs=load_dataset(project_key,"audit",["_id","Tarih","Kullanıcı","İşlem","Veri Seti","Kayıt Sayısı","Detay"])
    if not logs.empty: st.dataframe(logs.drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page == "Excel Raporları":
    st.title("Excel Raporları")
    report_tables={
        "Proje_Ozeti":pd.DataFrame([["Proje",st.session_state.current_project],["Sürüm",APP_VERSION],["Rapor Tarihi",report_date.isoformat()],["Para Birimi",currency],["Toplam Bütçe",budget],["Gerçekleşen Maliyet",actual_cost],["Fiziksel İlerleme (%)",progress],["BAC",evm["BAC"]],["PV",evm["PV"]],["EV",evm["EV"]],["AC",evm["AC"]],["CPI",evm["CPI"]],["SPI",evm["SPI"]],["EAC",evm["EAC"]]],columns=["Gösterge","Değer"]),
        "Butce_Metraj":plan.drop(columns=["_id"],errors="ignore"),"Gerceklesen":actual.drop(columns=["_id"],errors="ignore"),
        "Kategori_Sapma":cost_summary,"Is_Programi":schedule.drop(columns=["_id"],errors="ignore"),"CPM":cpm.drop(columns=["_id"],errors="ignore"),
        "Riskler":risk.drop(columns=["_id"],errors="ignore"),"Birim_Fiyat":DATA["library"].drop(columns=["_id"],errors="ignore"),
        "Satin_Alma":DATA["purchase"].drop(columns=["_id"],errors="ignore"),"Puantaj":DATA["labor"].drop(columns=["_id"],errors="ignore"),
        "Depo_Stok":DATA["inventory"].drop(columns=["_id"],errors="ignore"),"Hakedis":DATA["hakediş"].drop(columns=["_id"],errors="ignore"),
        "Nakit_Akisi":DATA["cash"].drop(columns=["_id"],errors="ignore"),
        "Sozlesmeler":DATA["contracts"].drop(columns=["_id"],errors="ignore"),"Degisiklikler":DATA["changes"].drop(columns=["_id"],errors="ignore"),
        "Isveren_Hakedis":DATA["employer_hakediş"].drop(columns=["_id"],errors="ignore"),"Gunluk_Santiye":DATA["daily_reports"].drop(columns=["_id"],errors="ignore"),
        "Onay_Gecmisi":DATA["approvals"].drop(columns=["_id"],errors="ignore")}
    xlsx=excel_export(report_tables)
    st.download_button("📥 Excel Rapor İndir",xlsx,file_name=f"Emir_Asel_V13_{now_tr():%Y%m%d_%H%M}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")

st.divider()
st.caption(f"{APP_NAME} | V{APP_VERSION} | {DB_KIND.upper()} | Karar destek amaçlıdır.")
