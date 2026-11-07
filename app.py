# -*- coding: utf-8 -*-
"""
EMİR ASEL İNŞAAT | ENTEGRE PROJE KONTROL VE MALİYET YÖNETİM SİSTEMİ V12.0

Kurulum:
    pip install -r requirements_Emir_Asel_V12.txt
Çalıştırma:
    streamlit run Emir_Asel_Proje_Kontrol_V12.py

Önerilen requirements.txt:
    streamlit>=1.36
    pandas>=2.0
    numpy>=1.24
    plotly>=5.20
    openpyxl>=3.1
    sqlalchemy>=2.0
    psycopg[binary]>=3.1       # Yalnızca Neon/PostgreSQL kullanılacaksa
    pypdf>=4.0                 # PDF metin yedek okuyucu
    acikpoz                    # Kuruluysa ÇŞİDB poz kataloğu ayrıştırma desteği

VERİTABANI:
- Varsayılan: uygulama klasöründe emir_asel_v12.db (SQLite)
- Bulut/Neon: DATABASE_URL ortam değişkeni veya Streamlit secrets içinde DATABASE_URL
  Örnek: postgresql+psycopg://USER:PASSWORD@HOST/DBNAME?sslmode=require
- Uygulama tek başına SQLite ile çalışır; Neon bağlantısı zorunlu değildir.

ÖNEMLİ:
- Bu araç yönetim ve karar destek amaçlıdır. Resmi hakediş, muhasebe, sözleşme,
  iş güvenliği veya mühendislik onayının yerine geçmez.
- Birim fiyat PDF ayrıştırma doğruluğu PDF'nin yapısına ve ayrıştırıcıya bağlıdır.
  İçe aktarılan pozlar ve fiyatlar yayın kaynağıyla karşılaştırılarak onaylanmalıdır.
- EVM için programdaki WBS'lerin bütçe WBS'leriyle eşleşmesi gerekir.
- Para birimi proje bazındadır. Farklı para birimleri aynı toplamda birleştirilmemelidir.
"""
from __future__ import annotations

import os
import re
import json
import uuid
import tempfile
import hashlib
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
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_VERSION = "12.0"
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
PURCHASE_COLS = ["_id","Sipariş No","WBS","Poz No","Malzeme / Hizmet","Tedarikçi","Sipariş Tarihi","Termin Tarihi","Sipariş Miktarı","Teslim Miktarı","Birim","Birim Fiyat","Sipariş Tutarı","Faturalanan Tutar","Ödenen Tutar","Durum","Not"]
LABOR_COLS = ["_id","Tarih","Ekip","İş Tanımı","Çalışan Sayısı","Gün","Kişi Günlük Ücret","Toplam İşçilik","WBS","Taşeron / Firma","Not"]
INVENTORY_COLS = ["_id","Tarih","Depo","Malzeme Kodu","Malzeme","Hareket Türü","Miktar","Birim","Birim Fiyat","Tutar","WBS","Belge No","Açıklama"]
HAKEDIS_COLS = ["_id","Hakediş No","Dönem","Tarih","Taşeron / Firma","WBS","Brüt Tutar","Kesinti","Net Hakediş","KDV Oranı","KDV Tutarı","Ödenecek Tutar","Ödenen","Durum","Açıklama"]
PROJECT_COLS = ["_id","Proje Adı","İşveren","Lokasyon","Sözleşme No","Sözleşme Tarihi","Başlangıç Tarihi","Bitiş Tarihi","Para Birimi","KDV Oranı","Risk Rezervi (%)","Proje Yöneticisi","Açıklama"]
LEGACY_CASH_COLS = ["_id","Dönem","Planlanan","Gerçekleşen","Net","Kümülatif","Kaynak"]
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
}
DATE_FIELDS = {
    "plan": ["Başlangıç", "Bitiş"], "actual": ["Tarih"],
    "schedule": ["Başlangıç", "Bitiş", "Fiili Başlangıç", "Fiili Bitiş"],
    "risk": ["Hedef Tarih"], "cash": ["Tarih", "Vade Tarihi"],
    "purchase": ["Sipariş Tarihi", "Termin Tarihi"], "labor": ["Tarih"],
    "inventory": ["Tarih"], "hakediş": ["Tarih"],
    "library": ["Yayın / Güncelleme Tarihi"],
    "project": ["Sözleşme Tarihi", "Başlangıç Tarihi", "Bitiş Tarihi"],
}

# ------------------------------ GENEL YARDIMCILAR ------------------------------
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
    """JSON-safe recursive conversion; NaN/NaT become null."""
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


def json_default(obj: Any):
    return json_safe_value(obj)


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

# ------------------------------ CPM MOTORU ------------------------------
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
        warnings.append("Tekrarlanan WBS kodu var. Her aktivite için benzersiz WBS kullanın; CPM sonuçları güvenilir olmayabilir.")
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
            warnings.append(f"{task}: geçersiz ilişki tipi FS olarak kabul edildi.")
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
        warnings.append("Döngüsel öncül ilişkisi tespit edildi. CPM hesaplanamadı; WBS/öncül bağlantılarını kontrol edin.")
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

# ------------------------------ EVM MOTORU ------------------------------
def calculate_evm(plan: pd.DataFrame, actual: pd.DataFrame, schedule: pd.DataFrame, report_date: date) -> dict[str, float]:
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
                if pd.isna(start) or pd.isna(end):
                    continue
                if end < start:
                    continue
                rd = pd.Timestamp(report_date)
                if rd < start:
                    fraction = 0.0
                elif rd >= end:
                    fraction = 1.0
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

# ------------------------------ VERİTABANI ------------------------------
def resolve_database_url() -> str:
    url = os.getenv("DATABASE_URL", "").strip()
    if not url:
        try:
            url = str(st.secrets.get("DATABASE_URL", "")).strip()
        except Exception:
            url = ""
    if not url:
        return f"sqlite:///{DB_FILE.as_posix()}"
    # Neon'ın standart postgresql:// URI'sini requirements dosyasındaki psycopg v3 sürücüsüne yönlendir.
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
        con.execute(text("CREATE INDEX IF NOT EXISTS ix_ea_records_project_dataset ON ea_records(project_key, dataset)"))


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
    record = {"_id": uid(), "Tarih": now_tr().isoformat(timespec="seconds"), "Kullanıcı": os.getenv("USER", "uygulama"),
              "İşlem": action, "Veri Seti": dataset, "Kayıt Sayısı": int(count), "Detay": detail}
    payload = to_payload(record)
    with ENGINE.begin() as con:
        con.execute(text("INSERT INTO ea_records(project_key,dataset,record_id,payload,updated_at) VALUES(:p,'audit',:r,:v,:t)"),
                    {"p": project_key, "r": record["_id"], "v": payload, "t": record["Tarih"]})


def save_dataset(project_key: str, dataset: str, df: pd.DataFrame, columns: list[str], make_audit: bool = True) -> None:
    """Dataset bazında transaction + record_id upsert. Diğer veri setleri silinmez."""
    frame = ensure_frame(df, columns)
    now_iso = now_tr().isoformat(timespec="seconds")
    statement = text("""
        INSERT INTO ea_records(project_key,dataset,record_id,payload,updated_at)
        VALUES(:p,:d,:r,:v,:t)
        ON CONFLICT(project_key,dataset,record_id)
        DO UPDATE SET payload=excluded.payload, updated_at=excluded.updated_at
    """)
    with ENGINE.begin() as con:
        current = con.execute(text("SELECT record_id FROM ea_records WHERE project_key=:p AND dataset=:d"), {"p": project_key, "d": dataset}).fetchall()
        current_ids = {str(row[0]) for row in current}
        keep_ids = set(frame["_id"].astype(str).tolist()) if "_id" in frame.columns else set()
        # Satır silme işlemi sadece ilgili proje + veri seti ile sınırlıdır.
        for stale_id in current_ids - keep_ids:
            con.execute(text("DELETE FROM ea_records WHERE project_key=:p AND dataset=:d AND record_id=:r"),
                        {"p": project_key, "d": dataset, "r": stale_id})
        for _, row in frame.iterrows():
            con.execute(statement, {"p": project_key, "d": dataset, "r": str(row["_id"]), "v": to_payload(row.to_dict()), "t": now_iso})
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
    """Yeni proje adı kaydedilirken taslak projedeki kayıtları tek transaction ile taşır."""
    if old_key == new_key:
        return
    with ENGINE.begin() as con:
        rows = con.execute(text("SELECT dataset, record_id FROM ea_records WHERE project_key=:old"), {"old": old_key}).fetchall()
        for dataset, record_id in rows:
            exists = con.execute(text("SELECT 1 FROM ea_records WHERE project_key=:new AND dataset=:d AND record_id=:r"),
                                 {"new": new_key, "d": dataset, "r": record_id}).first()
            if exists:
                # Hedefte aynı UUID varsa kaynak satır korunur; normal UUID akışında beklenmez.
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
    """Önce acikpoz; paket yoksa pypdf metin çıkarımı. Sonuç elle doğrulanmalıdır."""
    data = upload.getvalue()
    if len(data) > 25 * 1024 * 1024:
        raise ValueError("PDF 25 MB sınırını aşıyor. Dosyayı parçalara ayırıp yeniden yükleyin.")
    temp_path = None
    rows = []
    parser_used = ""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(data)
            temp_path = tmp.name
        try:
            from acikpoz import parse_catalog  # type: ignore
            parsed = parse_catalog(temp_path)
            for p in getattr(parsed, "pozes", []):
                if getattr(p, "is_priced", False) and getattr(p, "fiyat", None) is not None:
                    rows.append({"Poz No": str(p.poz_no).strip(), "İş Kalemi": str(p.tanim).strip()[:500],
                                 "Birim": str(p.birim).strip(), "Birim Fiyat": safe_num(p.fiyat),
                                 "Para Birimi": currency, "Fiyat Yılı": int(price_year)})
            parser_used = "acikpoz"
        except ImportError:
            try:
                from pypdf import PdfReader  # type: ignore
                reader = PdfReader(temp_path)
                text_blob = "\n".join(page.extract_text() or "" for page in reader.pages)
                # Genel metin yakalayıcı: poz kodu + açıklama + birim + fiyat aynı satırdaysa aday üretir.
                pattern = re.compile(r"(?m)^\s*([A-ZÇĞİÖŞÜa-zçğıöşü0-9.\-/]{4,})\s+(.{5,}?)\s+(Ad| adet|m|m2|m²|m3|m³|kg|ton|saat|gün|takım|set|%)\s+([\d.,]+)\s*$", re.I)
                for match in pattern.finditer(text_blob):
                    rows.append({"Poz No": match.group(1).strip(), "İş Kalemi": match.group(2).strip(),
                                 "Birim": match.group(3).strip(), "Birim Fiyat": safe_num(match.group(4)),
                                 "Para Birimi": currency, "Fiyat Yılı": int(price_year)})
                parser_used = "pypdf satır eşleştirme (aday kayıt)"
            except ImportError:
                raise RuntimeError("PDF okumak için 'acikpoz' veya 'pypdf' kurulmalı. Alternatif olarak Excel/CSV şablonunu içe aktarın.")
        if not rows:
            raise ValueError("PDF içinden güvenilir fiyatlı poz çıkarılamadı. PDF metin tabanlı olmayabilir veya tablo biçimi desteklenmiyor olabilir.")
        df = pd.DataFrame(rows)
        df = df[df["Poz No"].astype(str).str.strip().ne("")]
        df = df.drop_duplicates(subset=["Poz No", "Fiyat Yılı"], keep="last")
        return df, parser_used
    finally:
        if temp_path and os.path.exists(temp_path):
            os.unlink(temp_path)

def migrate_v11_legacy_tables() -> list[str]:
    """V11 Neon tablolarını yeni kayıtlara kopyalar; kaynak tabloları asla silmez."""
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

    def rename_columns(raw: pd.DataFrame, mapping: dict[str, str], cols: list[str]) -> pd.DataFrame:
        raw = raw.drop(columns=["id"], errors="ignore").copy()
        out = pd.DataFrame()
        for target, old in mapping.items():
            out[target] = raw[old] if old in raw.columns else None
        return ensure_frame(out, cols)

    migrations = {
        "plan": (PLAN_COLS, {
            "WBS":"wbs", "Poz No":"poz", "İş Kalemi":"is_kalemi", "Ana Kategori":"ana_kategori", "Alt Kategori":"alt_kategori",
            "Birim":"birim", "Miktar":"miktar", "Birim Fiyat":"birim_fiyat", "Planlanan Tutar":"planlanan_tutar",
            "Başlangıç":"baslangic", "Bitiş":"bitis", "Sorumlu":"sorumlu", "Açıklama":"aciklama"}),
        "actual": (ACTUAL_COLS, {
            "WBS":"wbs", "Poz No":"poz", "İş Kalemi":"is_kalemi", "Ana Kategori":"ana_kategori", "Birim":"birim",
            "Gerçekleşen Miktar":"gerceklesen_miktar", "Gerçekleşen Birim Fiyat":"gerceklesen_birim_fiyat",
            "Gerçekleşen Tutar":"gerceklesen_tutar", "Tarih":"tarih", "Tedarikçi":"tedarikci", "Belge No":"belge_no", "Açıklama":"aciklama"}),
        "schedule": (SCHEDULE_COLS, {
            "WBS":"wbs", "Aktivite":"aktivite", "Süre (gün)":"sure", "Öncül WBS":"oncul", "Bağımlılık Tipi":"bagimlilik",
            "Başlangıç":"baslangic", "Bitiş":"bitis", "Fiili Başlangıç":"fiili_baslangic", "Fiili Bitiş":"fiili_bitis",
            "Tamamlanma (%)":"tamamlanma", "Ağırlık (%)":"agirlik", "Sorumlu":"sorumlu", "Durum":"durum", "Kanıt / Saha Notu":"kanit"}),
        "risk": (RISK_COLS, {
            "Risk ID":"risk_id", "Risk Tanımı":"risk_tanimi", "Kategori":"kategori", "Olasılık (1-5)":"olasilik",
            "Etki (1-5)":"etki", "Risk Skoru":"risk_skoru", "Önlem / Aksiyon":"onlem", "Sorumlu":"sorumlu",
            "Durum":"durum", "Tahmini Maliyet Etkisi":"maliyet_etkisi"}),
    }
    migrated_any = False
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
            frame = rename_columns(raw, mapping, cols)
            frame["_id"] = [uid() for _ in range(len(frame))]
            if dataset == "plan":
                frame["Miktar"] = frame["Miktar"].apply(safe_num)
                frame["Birim Fiyat"] = frame["Birim Fiyat"].apply(safe_num)
                frame = calc_plan_amount(frame)
            elif dataset == "actual":
                frame["Gerçekleşen Miktar"] = frame["Gerçekleşen Miktar"].apply(safe_num)
                frame["Gerçekleşen Birim Fiyat"] = frame["Gerçekleşen Birim Fiyat"].apply(safe_num)
                frame = calc_actual_amount(frame)
            elif dataset == "schedule":
                frame["Süre (gün)"] = frame["Süre (gün)"].apply(safe_num)
                frame["Tamamlanma (%)"] = frame["Tamamlanma (%)"].apply(safe_num).clip(0,100)
                frame["Ağırlık (%)"] = frame["Ağırlık (%)"].apply(safe_num)
                frame["Bağımlılık Tipi"] = frame["Bağımlılık Tipi"].fillna("FS").astype(str).str.upper()
            elif dataset == "risk":
                frame["Olasılık (1-5)"] = frame["Olasılık (1-5)"].apply(safe_num)
                frame["Etki (1-5)"] = frame["Etki (1-5)"].apply(safe_num)
                frame["Risk Skoru"] = frame["Olasılık (1-5)"] * frame["Etki (1-5)"]
            save_dataset(target_project, dataset, frame, cols)
            notes.append(f"V11 {dataset} tablosundan {len(frame)} kayıt '{target_project}' projesine kopyalandı.")
            migrated_any = True
        except Exception as exc:
            notes.append(f"V11 {dataset} tablosu aktarılamadı: {exc}")

    if "cash" in existing_tables:
        try:
            current_cash = load_dataset(target_project, "legacy_cash", LEGACY_CASH_COLS)
            if current_cash.empty:
                raw_cash = read_legacy("cash")
                if not raw_cash.empty:
                    old_cash = pd.DataFrame({
                        "Dönem": raw_cash["donem"] if "donem" in raw_cash.columns else None,
                        "Planlanan": raw_cash["planlanan"].apply(safe_num) if "planlanan" in raw_cash.columns else 0,
                        "Gerçekleşen": raw_cash["gerceklesen"].apply(safe_num) if "gerceklesen" in raw_cash.columns else 0,
                        "Net": raw_cash["net"].apply(safe_num) if "net" in raw_cash.columns else 0,
                        "Kümülatif": raw_cash["kumulatif"].apply(safe_num) if "kumulatif" in raw_cash.columns else 0,
                        "Kaynak": "V11 cash tablosu — arşiv kopyası"})
                    old_cash = ensure_frame(old_cash, LEGACY_CASH_COLS)
                    old_cash["_id"] = [uid() for _ in range(len(old_cash))]
                    save_dataset(target_project, "legacy_cash", old_cash, LEGACY_CASH_COLS)
                    notes.append(f"V11 nakit akışı tablosundan {len(old_cash)} dönem arşivlendi.")
                    migrated_any = True
        except Exception as exc:
            notes.append(f"V11 nakit akışı tablosu arşivlenemedi: {exc}")

    if "library" in existing_tables:
        try:
            current_library = load_dataset("__GLOBAL__", "library", LIB_COLS)
            if current_library.empty:
                raw_lib = read_legacy("library")
                if not raw_lib.empty:
                    old_lib = pd.DataFrame({
                        "Poz No": raw_lib["poz"].astype(str) if "poz" in raw_lib.columns else "",
                        "İş Kalemi": raw_lib["is_kalemi"].astype(str) if "is_kalemi" in raw_lib.columns else "",
                        "Birim": raw_lib["birim"].astype(str) if "birim" in raw_lib.columns else "",
                        "Birim Fiyat": raw_lib["birim_fiyat"].apply(safe_num) if "birim_fiyat" in raw_lib.columns else 0,
                        "Para Birimi": "Belirsiz", "Fiyat Yılı": 0,
                        "Kaynak": "V11 migrasyonu — fiyat yılı/para birimi/kaynak doğrulanmalı",
                        "Kaynak Dosya": "", "Yayın / Güncelleme Tarihi": now_tr().date().isoformat(), "Aktif": True})
                    old_lib = ensure_frame(old_lib, LIB_COLS)
                    old_lib["_id"] = [uid() for _ in range(len(old_lib))]
                    save_dataset("__GLOBAL__", "library", old_lib, LIB_COLS)
                    notes.append(f"V11 birim fiyat kütüphanesinden {len(old_lib)} kayıt kopyalandı; yıl ve para birimi bilgileri doğrulanmalı.")
                    migrated_any = True
        except Exception as exc:
            notes.append(f"V11 birim fiyat kütüphanesi aktarılamadı: {exc}")

    if migrated_any:
        projects = load_dataset("__GLOBAL__", "project", PROJECT_COLS)
        if not projects["Proje Adı"].astype(str).eq(target_project).any():
            meta = pd.DataFrame([{"_id":uid(),"Proje Adı":target_project,"İşveren":"","Lokasyon":"","Sözleşme No":"",
                "Sözleşme Tarihi":None,"Başlangıç Tarihi":None,"Bitiş Tarihi":None,"Para Birimi":"₺","KDV Oranı":0.20,
                "Risk Rezervi (%)":0.05,"Proje Yöneticisi":"","Açıklama":"V11.0 verilerinden otomatik ve silmeden aktarılmış proje."}])
            projects = pd.concat([projects,meta],ignore_index=True)
            save_dataset("__GLOBAL__","project",projects,PROJECT_COLS)
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
    add("Bütçede mükerrer poz/WBS", plan.duplicated(subset=["WBS", "Poz No"], keep=False).sum(), "Aynı WBS altında farklı pozlar olabilir; aynı WBS + poz çiftini kontrol edin.")
    add("Negatif metraj", (pd.to_numeric(plan["Miktar"], errors="coerce") < 0).sum(), "Miktar negatif olmamalıdır.")
    add("Negatif birim fiyat", (pd.to_numeric(plan["Birim Fiyat"], errors="coerce") < 0).sum(), "Birim fiyat negatif olmamalıdır.")
    pstart, pend = pd.to_datetime(plan["Başlangıç"], errors="coerce", format="mixed", dayfirst=True), pd.to_datetime(plan["Bitiş"], errors="coerce", format="mixed", dayfirst=True)
    add("Bitiş tarihi başlangıçtan önce", ((pend < pstart) & pend.notna() & pstart.notna()).sum(), "Bütçe satırlarının tarih sıralamasını kontrol edin.")
    add("WBS'siz gerçekleşen", actual["WBS"].fillna("").astype(str).str.strip().eq("").sum(), "Fatura/gerçekleşen kaydını bütçe WBS koduna bağlayın.")
    known = set(plan["WBS"].dropna().astype(str).str.strip())
    add("Bütçede bulunmayan gerçekleşen WBS", (~actual["WBS"].fillna("").astype(str).str.strip().isin(known)).sum(), "Bütçeye eklenmemiş gerçekleşen iş veya hatalı WBS olabilir.")
    pr = pd.to_numeric(schedule["Tamamlanma (%)"], errors="coerce")
    add("İlerleme %0–100 dışında", ((pr < 0) | (pr > 100)).sum(), "Aktivite ilerlemesi yüzde 0–100 aralığında olmalıdır.")
    add("Tekrarlanan aktivite WBS", schedule["WBS"].dropna().astype(str).duplicated().sum(), "CPM hesabı için aktivite WBS kodları benzersiz olmalıdır.")
    add("Risk puanı eksik", (risk["Risk Skoru"].isna() | risk["Risk Skoru"].eq(0)).sum(), "Olasılık ve etki puanlarını girin.")
    add("Kütüphane poz/fiyat boş", (library["Poz No"].fillna("").astype(str).str.strip().eq("") | pd.to_numeric(library["Birim Fiyat"], errors="coerce").isna()).sum(), "Ana fiyat kütüphanesindeki kayıtları doğrulayın.")
    return pd.DataFrame(checks)

# ------------------------------ VERİTABANI DURUMU ------------------------------
if DB_ERROR:
    st.error(f"Veritabanı başlatılamadı: {DB_ERROR}")
    st.stop()
else:
    LEGACY_MIGRATION_NOTES = migrate_v11_legacy_tables()
    for migration_note in LEGACY_MIGRATION_NOTES:
        st.info(migration_note)

# ------------------------------ PROJE SEÇİMİ ------------------------------
known_projects = load_dataset("__GLOBAL__", "project", PROJECT_COLS)
project_names = sorted(set(known_projects["Proje Adı"].dropna().astype(str).str.strip()) - {""}) if not known_projects.empty else []
default_project = "Yeni Proje"
if "current_project" not in st.session_state:
    st.session_state.current_project = project_names[0] if project_names else default_project

st.markdown(f'<div class="ea-header"><h1>🏗️ {APP_NAME}</h1><div class="sub">ENTEGRE PROJE KONTROL · MALİYET · METRAJ · CPM · EVM · TEDARİK · DEPO · HAKEDİŞ · V{APP_VERSION} · {now_tr():%d.%m.%Y %H:%M}</div></div>', unsafe_allow_html=True)

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
    MENU = ["Yönetici Dashboard", "Proje Künyesi / İçe Aktarma", "Birim Fiyat Kütüphanesi", "Bütçe ve Metraj", "Metraj Kontrolü", "Gerçekleşen Maliyet", "İş Programı / CPM", "EVM Performans", "Satın Alma ve Tedarik", "İşçilik ve Puantaj", "Depo / Stok", "Hakediş Takibi", "Nakit Akışı / S-Curve", "Risk Yönetimi", "Veri Kontrolü", "Yedekleme / Denetim", "Excel Raporları"]
    page = st.radio("MODÜLLER", MENU)
    st.divider()
    st.caption(f"Veritabanı: {'Neon / PostgreSQL' if DB_KIND == 'postgresql' else 'Yerel SQLite'}")

project_key = st.session_state.current_project
if project_key == default_project:
    project_key = "__DRAFT_PROJECT__"

# Load all active project data (cached in session by project key)
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

# ------------------------------ KAYDETME / STATE YARDIMCILARI ------------------------------
def save_from_editor(dataset: str, edited: pd.DataFrame, columns: list[str], recalc: str | None = None):
    frame = ensure_frame(edited, columns)
    if recalc == "plan": frame = calc_plan_amount(frame)
    elif recalc == "actual": frame = calc_actual_amount(frame)
    elif recalc == "risk":
        frame["Risk Skoru"] = pd.to_numeric(frame["Olasılık (1-5)"], errors="coerce").fillna(0) * pd.to_numeric(frame["Etki (1-5)"], errors="coerce").fillna(0)
    elif recalc == "purchase":
        frame["Sipariş Tutarı"] = frame["Sipariş Miktarı"].apply(safe_num) * frame["Birim Fiyat"].apply(safe_num)
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
    st.success(f"{len(frame)} kayıt {dataset} veri setine güvenli şekilde kaydedildi.")
    st.rerun()


def reset_dataset_editor(dataset: str):
    st.session_state.pop(f"{dataset}_editor", None)
    st.rerun()

# ------------------------------ SAYFA: DASHBOARD ------------------------------
if page == "Yönetici Dashboard":
    st.title(f"Yönetici Dashboard | {st.session_state.current_project}")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Toplam Bütçe", money(budget, currency))
    c2.metric("Gerçekleşen Maliyet", money(actual_cost, currency), delta=money(actual_cost-budget, currency) if budget else None, delta_color="inverse")
    c3.metric("Fiziksel İlerleme", pct(progress))
    c4.metric("Tahmini Bütçe (EAC)", money(evm["EAC"], currency) if pd.notna(evm["EAC"]) else "—")
    c5,c6,c7,c8 = st.columns(4)
    c5.metric("Kalan Bütçe", money(budget-actual_cost, currency))
    c6.metric("CPI", f'{evm["CPI"]:.2f}' if pd.notna(evm["CPI"]) else "—")
    c7.metric("SPI", f'{evm["SPI"]:.2f}' if pd.notna(evm["SPI"]) else "—")
    c8.metric("Açık Risk", int(risk["Durum"].astype(str).str.casefold().isin(["açık","acik","open"]).sum()) if not risk.empty else 0)
    left, right = st.columns([1.35,1])
    with left:
        st.subheader("Kategori Bazında Bütçe / Gerçekleşen")
        if not cost_summary.empty:
            chart = cost_summary.melt(id_vars="Ana Kategori", value_vars=["Bütçe","Gerçekleşen"], var_name="Tür", value_name="Tutar")
            fig = px.bar(chart, x="Ana Kategori", y="Tutar", color="Tür", barmode="group", color_discrete_sequence=[NAVY,RED])
            fig.update_layout(xaxis_tickangle=-35, margin=dict(l=10,r=10,t=20,b=130))
            st.plotly_chart(chart_layout(fig, 440), use_container_width=True)
        else: st.info("Henüz bütçe veya gerçekleşen maliyet kaydı yok.")
    with right:
        st.subheader("Bütçe Dağılımı")
        pie = cost_summary[cost_summary["Bütçe"] > 0] if not cost_summary.empty else pd.DataFrame()
        if not pie.empty:
            fig = px.pie(pie, names="Ana Kategori", values="Bütçe", hole=.58, color_discrete_sequence=PALETTE)
            st.plotly_chart(chart_layout(fig, 440), use_container_width=True)
        else: st.info("Bütçe dağılımı için veri girilmedi.")
    st.subheader("Proje Sağlık Göstergeleri")
    g1,g2 = st.columns(2)
    with g1: st.plotly_chart(gauge_chart(evm["CPI"], "CPI — Maliyet Performansı"), use_container_width=True)
    with g2: st.plotly_chart(gauge_chart(evm["SPI"], "SPI — Zaman Performansı"), use_container_width=True)
    st.subheader("En Yüksek Bütçeli 15 İş Kalemi")
    if not plan.empty:
        st.dataframe(plan.nlargest(15,"Planlanan Tutar")[["WBS","Poz No","İş Kalemi","Ana Kategori","Miktar","Birim","Birim Fiyat","Planlanan Tutar"]], use_container_width=True, hide_index=True)
    st.caption("CPI/SPI sonuçları, bütçe ve iş programındaki WBS eşleşmesine ve raporlama tarihine bağlıdır.")

# ------------------------------ PROJE KÜNYESİ / İÇE AKTARMA ------------------------------
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
        contract = d.text_input("Sözleşme No", value=default_meta.get("Sözleşme No", ""))
        e,f = st.columns(2)
        contract_date = e.date_input("Sözleşme Tarihi", value=pd.to_datetime(default_meta.get("Sözleşme Tarihi"), errors="coerce", format="mixed", dayfirst=True).date() if pd.notna(pd.to_datetime(default_meta.get("Sözleşme Tarihi"), errors="coerce", format="mixed", dayfirst=True)) else now_tr().date())
        start_date = f.date_input("Proje Başlangıcı", value=pd.to_datetime(default_meta.get("Başlangıç Tarihi"), errors="coerce", format="mixed", dayfirst=True).date() if pd.notna(pd.to_datetime(default_meta.get("Başlangıç Tarihi"), errors="coerce", format="mixed", dayfirst=True)) else now_tr().date())
        g,h = st.columns(2)
        end_date = g.date_input("Planlanan Bitiş", value=pd.to_datetime(default_meta.get("Bitiş Tarihi"), errors="coerce", format="mixed", dayfirst=True).date() if pd.notna(pd.to_datetime(default_meta.get("Bitiş Tarihi"), errors="coerce", format="mixed", dayfirst=True)) else now_tr().date()+timedelta(days=365))
        meta_currency = h.selectbox("Proje Para Birimi", ["₺","$","€"], index=["₺","$","€"].index(default_meta.get("Para Birimi","₺")) if default_meta.get("Para Birimi","₺") in ["₺","$","€"] else 0)
        i,j = st.columns(2)
        vat = i.number_input("KDV Oranı", 0.0, 1.0, safe_num(default_meta.get("KDV Oranı",0.20)), 0.01)
        reserve = j.number_input("Risk Rezervi (%)", 0.0, 1.0, safe_num(default_meta.get("Risk Rezervi (%)",0.05)), 0.01)
        manager = st.text_input("Proje Yöneticisi", value=default_meta.get("Proje Yöneticisi", ""))
        note = st.text_area("Açıklama", value=default_meta.get("Açıklama", ""))
        submitted = st.form_submit_button("Proje Künyesini Kaydet", type="primary")
        if submitted:
            if not pname.strip(): st.error("Proje adı boş bırakılamaz.")
            else:
                meta = pd.DataFrame([{ "_id": str(default_meta.get("_id") or uid()), "Proje Adı": pname.strip(), "İşveren":owner, "Lokasyon":location,
                    "Sözleşme No":contract, "Sözleşme Tarihi":contract_date, "Başlangıç Tarihi":start_date, "Bitiş Tarihi":end_date,
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
                st.success("Proje künyesi kaydedildi. Yeni proje adıyla devam ediliyor.")
                st.rerun()
    st.divider()
    st.subheader("Mevcut Excel Dosyasından İçe Aktarma")
    uploaded = st.file_uploader("Excel dosyası yükle (.xlsx / .xlsm)", type=["xlsx","xlsm"], key="project_import_excel")
    if uploaded:
        try:
            sheets = pd.read_excel(uploaded, sheet_name=None)
            sheet_names = list(sheets.keys())
            st.write("Bulunan sayfalar:", sheet_names)
            plan_sheet = st.selectbox("Bütçe / metraj sayfası", sheet_names)
            actual_sheet = st.selectbox("Gerçekleşen maliyet sayfası", ["—"] + sheet_names)
            schedule_sheet = st.selectbox("İş programı sayfası", ["—"] + sheet_names)
            if st.button("Seçilen sayfaları içe aktar", type="primary"):
                p = map_imported_sheet(sheets[plan_sheet], "plan")
                save_dataset(project_key,"plan",p,PLAN_COLS); st.session_state["data_plan"] = p
                if actual_sheet != "—":
                    a = map_imported_sheet(sheets[actual_sheet],"actual")
                    save_dataset(project_key,"actual",a,ACTUAL_COLS); st.session_state["data_actual"] = a
                if schedule_sheet != "—":
                    s = map_imported_sheet(sheets[schedule_sheet],"schedule")
                    save_dataset(project_key,"schedule",s,SCHEDULE_COLS); st.session_state["data_schedule"] = s
                audit_log(project_key,"İÇE AKTAR", "excel", len(p), uploaded.name)
                st.success("İçe aktarma tamamlandı. Kayıtları ilgili modüllerde kontrol edin.")
                st.rerun()
        except Exception as exc: st.error(f"Excel içe aktarma hatası: {exc}")
    st.subheader("Boş Excel Veri Şablonu")
    st.download_button("Şablonları indir", excel_export({"Butce_Metraj":pd.DataFrame(columns=PLAN_COLS),"Gerceklesen":pd.DataFrame(columns=ACTUAL_COLS),"Is_Programi":pd.DataFrame(columns=SCHEDULE_COLS),"Riskler":pd.DataFrame(columns=RISK_COLS)}), file_name="Emir_Asel_V12_Sablon.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ------------------------------ BİRİM FİYAT KÜTÜPHANESİ ------------------------------
elif page == "Birim Fiyat Kütüphanesi":
    st.title("Merkezi Birim Fiyat Kütüphanesi")
    st.markdown('<div class="ea-note">PDF fiyat listeleri merkezi kütüphaneye kaydedilir. Fiyat yılı ve kaynak bilgisi korunur; eski yılların fiyatları silinmeden saklanır.</div>', unsafe_allow_html=True)
    price_year = st.number_input("PDF / Fiyat Yılı", min_value=2000, max_value=2100, value=now_tr().year, step=1)
    pdf_currency = st.selectbox("PDF Para Birimi", ["₺","$","€"], key="pdf_currency")
    pdf_upload = st.file_uploader("Birim fiyat PDF yükle (maksimum 25 MB)", type=["pdf"], key="price_pdf")
    if pdf_upload and st.button("PDF'i oku ve kütüphaneye aktar", type="primary"):
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
            # Aynı yıl + poz + kaynak dosya yinelenirse tekrar ekleme.
            keys = set(zip(existing["Poz No"].astype(str), existing["Fiyat Yılı"].astype(str), existing["Kaynak Dosya"].astype(str))) if not existing.empty else set()
            parsed = parsed[[ (str(r["Poz No"]), str(r["Fiyat Yılı"]), str(r["Kaynak Dosya"])) not in keys for _,r in parsed.iterrows() ]]
            combined = pd.concat([existing, parsed], ignore_index=True)
            save_dataset("__GLOBAL__","library",combined,LIB_COLS)
            st.session_state["data_library"] = ensure_frame(combined, LIB_COLS)
            st.success(f"{len(parsed)} fiyatlı poz işlendi. Ayrıştırıcı: {parser_name}. Fiyatları kaynak PDF ile doğrulayın.")
            st.rerun()
        except Exception as exc: st.error(f"PDF işleme hatası: {exc}")
    with st.expander("Manuel poz ekle / fiyat revizyonu", expanded=False):
        with st.form("manual_price_form"):
            a,b = st.columns(2)
            poz = a.text_input("Poz No")
            item = b.text_input("İş Kalemi")
            c,d,e = st.columns(3)
            unit = c.text_input("Birim", "Ad")
            price = d.number_input("Birim Fiyat", min_value=0.0, step=1.0, format="%.4f")
            year = e.number_input("Fiyat Yılı", min_value=2000, max_value=2100, value=now_tr().year)
            f,g = st.columns(2)
            pcurr = f.selectbox("Para Birimi", ["₺","$","€"], key="manual_currency")
            source = g.text_input("Kaynak / Yayın")
            if st.form_submit_button("Fiyatı kütüphaneye kaydet"):
                if not poz.strip() or not item.strip(): st.error("Poz numarası ve iş kalemi zorunludur.")
                else:
                    new = {"_id":uid(),"Poz No":poz.strip(),"İş Kalemi":item.strip(),"Birim":unit,"Birim Fiyat":price,"Para Birimi":pcurr,"Fiyat Yılı":int(year),"Kaynak":source or "Manuel giriş","Kaynak Dosya":"","Yayın / Güncelleme Tarihi":now_tr().date().isoformat(),"Aktif":True}
                    updated = pd.concat([DATA["library"],pd.DataFrame([new])],ignore_index=True)
                    save_dataset("__GLOBAL__","library",updated,LIB_COLS)
    price_table_upload = st.file_uploader("PDF dışı fiyat kütüphanesi içe aktar (Excel / CSV)", type=["xlsx","xlsm","csv"], key="price_table_import")
    if price_table_upload and st.button("Excel / CSV fiyat listesini içe aktar"):
        try:
            if price_table_upload.name.lower().endswith(".csv"):
                raw_prices = pd.read_csv(price_table_upload)
            else:
                raw_sheets = pd.read_excel(price_table_upload, sheet_name=None)
                sheet_choice = next(iter(raw_sheets))
                raw_prices = raw_sheets[sheet_choice]
                if len(raw_sheets) > 1:
                    st.info(f"İçe aktarmada ilk sayfa kullanıldı: {sheet_choice}. Diğer sayfaları ayrı dosya olarak yükleyin.")
            c_poz = find_column(raw_prices,["Poz No","Poz","Poz Numarası"])
            c_item = find_column(raw_prices,["İş Kalemi","Açıklama","Tanım","Poz Tanımı"])
            c_unit = find_column(raw_prices,["Birim","Ölçü Birimi"])
            c_price = find_column(raw_prices,["Birim Fiyat","Fiyat","BF"])
            c_year = find_column(raw_prices,["Fiyat Yılı","Yıl"])
            if not all([c_poz,c_item,c_unit,c_price]):
                raise ValueError("Dosyada Poz No, İş Kalemi, Birim ve Birim Fiyat kolonları bulunmalı.")
            parsed_prices=pd.DataFrame({"Poz No":raw_prices[c_poz].astype(str).str.strip(),"İş Kalemi":raw_prices[c_item].astype(str).str.strip(),"Birim":raw_prices[c_unit].astype(str).str.strip(),"Birim Fiyat":raw_prices[c_price].apply(safe_num),"Para Birimi":currency,"Fiyat Yılı":raw_prices[c_year].apply(safe_num).replace(0,now_tr().year) if c_year else int(now_tr().year),"Kaynak":f"Tablo: {price_table_upload.name}","Kaynak Dosya":price_table_upload.name,"Yayın / Güncelleme Tarihi":now_tr().date().isoformat(),"Aktif":True})
            parsed_prices=parsed_prices[parsed_prices["Poz No"].ne("") & parsed_prices["İş Kalemi"].ne("")]
            parsed_prices=ensure_frame(parsed_prices,LIB_COLS)
            merged=pd.concat([DATA["library"],parsed_prices],ignore_index=True)
            save_dataset("__GLOBAL__","library",merged,LIB_COLS)
            st.session_state["data_library"]=ensure_frame(merged,LIB_COLS)
            st.success(f"{len(parsed_prices)} satır kütüphaneye aktarıldı. Poz ve fiyatları doğrulayın."); st.rerun()
        except Exception as exc:
            st.error(f"Fiyat tablosu içe aktarılamadı: {exc}")
    lib = DATA["library"].copy()
    search = st.text_input("Poz No / İş Kalemi / Kaynak ara", key="library_search")
    if search:
        mask = lib.astype(str).apply(lambda c:c.str.contains(search,case=False,na=False)).any(axis=1)
        lib = lib[mask]
    st.dataframe(lib.drop(columns=["_id"], errors="ignore"), use_container_width=True, hide_index=True, height=430)
    if not lib.empty:
        st.download_button("Kütüphane Excel dışa aktar", excel_export({"Birim_Fiyat_Kutuphanesi":lib}), file_name="Emir_Asel_Birim_Fiyat_Kutuphanesi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ------------------------------ BÜTÇE VE METRAJ ------------------------------
elif page == "Bütçe ve Metraj":
    st.title("Proje Bütçesi ve Metraj Yönetimi")
    library = DATA["library"].copy()
    active_lib = library[library["Aktif"].astype(str).str.lower().isin(["true","1","evet"]) | library["Aktif"].isna()] if not library.empty else library
    if not active_lib.empty:
        with st.expander("Kütüphaneden poz seçerek bütçeye ekle", expanded=False):
            labels = active_lib.apply(lambda r:f"{r['Poz No']} | {r['İş Kalemi']} | {r['Fiyat Yılı']} | {r['Birim Fiyat']} {r['Para Birimi']}",axis=1).tolist()
            selected = st.selectbox("Poz seç", labels)
            selected_row = active_lib.iloc[labels.index(selected)] if selected in labels else None
            if selected_row is not None:
                st.write(f"**Birim:** {selected_row['Birim']} · **Kaynak:** {selected_row['Kaynak']}")
                if str(selected_row.get("Para Birimi", currency)) != currency:
                    st.warning(f"Kütüphane fiyatı {selected_row.get('Para Birimi')} cinsinden; aktif proje para birimi {currency}. Fiyatı bütçeye aktarmadan önce kur dönüşümünü yapın.")
                qty = st.number_input("Metraj / Miktar", min_value=0.0, step=1.0, key="library_add_qty")
                wbs = st.text_input("WBS", key="library_add_wbs")
                if st.button("Seçili pozu bütçeye ekle"):
                    new = {c:None for c in PLAN_COLS}
                    new.update({"_id":uid(),"WBS":wbs,"Poz No":selected_row["Poz No"],"İş Kalemi":selected_row["İş Kalemi"],"Ana Kategori":auto_categorize(selected_row["İş Kalemi"]),"Birim":selected_row["Birim"],"Miktar":qty,"Birim Fiyat":safe_num(selected_row["Birim Fiyat"]),"Fiyat Yılı":selected_row["Fiyat Yılı"],"Fiyat Kaynağı":selected_row["Kaynak"],"Metraj Durumu":"Taslak"})
                    added = pd.concat([DATA["plan"],pd.DataFrame([new])],ignore_index=True)
                    added = calc_plan_amount(ensure_frame(added,PLAN_COLS))
                    save_dataset(project_key,"plan",added,PLAN_COLS); st.session_state["data_plan"] = added
                    st.success("Poz bütçeye eklendi."); st.rerun()
    else: st.info("Birim fiyat kütüphanesi boş. Kütüphaneye PDF veya manuel poz ekleyebilirsiniz.")
    st.caption("Planlanan Tutar = Miktar × Birim Fiyat. Formül kaydetme aşamasında yeniden hesaplanır.")
    edited = st.data_editor(plan, num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={"_id":None,"Ana Kategori":st.column_config.SelectboxColumn(options=COST_CATEGORIES),
            "Metraj Durumu":st.column_config.SelectboxColumn(options=["Taslak","Kontrol Bekliyor","Onaylandı","Revize","Tahmini"]),
            "Miktar":st.column_config.NumberColumn(min_value=0,format="%.3f"),
            "Birim Fiyat":st.column_config.NumberColumn(min_value=0,format="%.4f"),
            "Planlanan Tutar":st.column_config.NumberColumn(disabled=True,format="%.2f"),
            "Başlangıç":st.column_config.DateColumn(),"Bitiş":st.column_config.DateColumn()}, key="plan_editor")
    if st.button("Bütçe / metraj değişikliklerini güvenli kaydet", type="primary"):
        save_from_editor("plan",edited,PLAN_COLS,"plan")
    if not cost_summary.empty:
        st.subheader("Kategori Özeti"); st.dataframe(cost_summary,use_container_width=True,hide_index=True)
    st.metric("Toplam Planlanan Bütçe",money(budget,currency))

# ------------------------------ METRAJ KONTROLÜ ------------------------------
elif page == "Metraj Kontrolü":
    st.title("Metraj Kontrol ve Onay Paneli")
    st.caption("Metraj durumları bütçe tablosunda düzenlenir. Bu ekran açık, tahmini, revize ve onaylı metrajları ayırır.")
    metraj_view = plan.copy()
    status_filter = st.multiselect("Metraj durumu", ["Taslak","Kontrol Bekliyor","Onaylandı","Revize","Tahmini"], default=["Taslak","Kontrol Bekliyor","Revize","Tahmini"])
    if status_filter:
        metraj_view = metraj_view[metraj_view["Metraj Durumu"].fillna("Taslak").isin(status_filter)]
    c1,c2,c3 = st.columns(3)
    pending = plan[plan["Metraj Durumu"].fillna("Taslak").isin(["Taslak","Kontrol Bekliyor","Revize","Tahmini"])]
    approved = plan[plan["Metraj Durumu"].fillna("").eq("Onaylandı")]
    c1.metric("Toplam Metraj Kalemi",len(plan)); c2.metric("Kontrol / Revize Bekleyen",len(pending)); c3.metric("Onaylı Kalem",len(approved))
    if not metraj_view.empty:
        st.dataframe(metraj_view.drop(columns=["_id"],errors="ignore"),use_container_width=True,hide_index=True)
    else:
        st.info("Seçilen durumda metraj kalemi bulunmuyor.")
    st.subheader("Metraj Durumu Dağılımı")
    if not plan.empty:
        stat = plan.assign(Metraj_Durumu=plan["Metraj Durumu"].fillna("Taslak")).groupby("Metraj_Durumu").agg(Kalem=("Poz No","count"), Tutar=("Planlanan Tutar","sum")).reset_index()
        st.dataframe(stat,use_container_width=True,hide_index=True)
        fig=px.pie(stat,names="Metraj_Durumu",values="Kalem",hole=.55,color_discrete_sequence=PALETTE)
        st.plotly_chart(chart_layout(fig,360),use_container_width=True)
    st.warning("Bu ekran metrajın teknik doğruluğunu onaylamaz. Onay statüsü yetkili teknik personel kontrolünden sonra verilmelidir.")


# ------------------------------ GERÇEKLEŞEN MALİYET ------------------------------
elif page == "Gerçekleşen Maliyet":
    st.title("Gerçekleşen Maliyet / Fatura / Hakediş Giderleri")
    st.info("Her maliyet kaydında WBS, belge/fatura no ve tedarikçi girin. Puantaj ve satın alma modüllerindeki tutarlar otomatik olarak buraya kopyalanmaz; çift kayıt engellenir.")
    edited = st.data_editor(actual,num_rows="dynamic",use_container_width=True,hide_index=True,
        column_config={"_id":None,"Ana Kategori":st.column_config.SelectboxColumn(options=COST_CATEGORIES),
            "Kayıt Türü":st.column_config.SelectboxColumn(options=["Fatura","Taşeron Hakedişi","Malzeme","Hizmet","Diğer"]),
            "Gerçekleşen Miktar":st.column_config.NumberColumn(min_value=0,format="%.3f"),
            "Gerçekleşen Birim Fiyat":st.column_config.NumberColumn(min_value=0,format="%.4f"),
            "Gerçekleşen Tutar":st.column_config.NumberColumn(disabled=True,format="%.2f"),"Tarih":st.column_config.DateColumn()}, key="actual_editor")
    if st.button("Gerçekleşen maliyetleri güvenli kaydet",type="primary"):
        save_from_editor("actual",edited,ACTUAL_COLS,"actual")
    st.subheader("Kategori Bazında Bütçe Sapması")
    st.dataframe(cost_summary,use_container_width=True,hide_index=True)
    over = cost_summary[cost_summary["Sapma"]>0] if not cost_summary.empty else pd.DataFrame()
    if not over.empty: st.warning("Bütçeyi aşan kategoriler: "+", ".join(over["Ana Kategori"].astype(str).tolist()))

# ------------------------------ İŞ PROGRAMI / CPM ------------------------------
elif page == "İş Programı / CPM":
    st.title("İş Programı, Gantt ve Kritik Yol")
    st.markdown('<div class="ea-note">Öncül alanı örnekleri: <b>A01</b> (FS varsayılan), <b>A01:FS</b>, <b>A01:SS:2</b>, <b>A01:FF;A02:SS</b>. İlişki türleri FS / SS / FF / SF; son sayı gün cinsinden gecikme/lead değeridir.</div>',unsafe_allow_html=True)
    edited = st.data_editor(schedule,num_rows="dynamic",use_container_width=True,hide_index=True,
        column_config={"_id":None,"Süre (gün)":st.column_config.NumberColumn(min_value=0,step=1),
            "Bağımlılık Tipi":st.column_config.SelectboxColumn(options=["FS","SS","FF","SF"]),
            "Tamamlanma (%)":st.column_config.NumberColumn(min_value=0,max_value=100,format="%.1f"),
            "Ağırlık (%)":st.column_config.NumberColumn(min_value=0,format="%.2f"),
            "Başlangıç":st.column_config.DateColumn(),"Bitiş":st.column_config.DateColumn(),
            "Fiili Başlangıç":st.column_config.DateColumn(),"Fiili Bitiş":st.column_config.DateColumn()},key="schedule_editor")
    if st.button("İş programını güvenli kaydet",type="primary"):
        frame=ensure_frame(edited,SCHEDULE_COLS)
        frame["Süre (gün)"]=pd.to_numeric(frame["Süre (gün)"],errors="coerce").fillna(0).clip(lower=0)
        frame["Tamamlanma (%)"]=pd.to_numeric(frame["Tamamlanma (%)"],errors="coerce").fillna(0).clip(0,100)
        frame["Ağırlık (%)"]=pd.to_numeric(frame["Ağırlık (%)"],errors="coerce").fillna(0).clip(lower=0)
        save_from_editor("schedule",frame,SCHEDULE_COLS)
    for w in cpm_warnings: st.warning(w)
    if not cpm.empty:
        st.subheader("CPM Sonuçları")
        st.metric("Hesaplanan Proje Süresi",f'{cpm["Erken Bitiş"].max():.0f} gün')
        st.dataframe(cpm.drop(columns=["_id"],errors="ignore"),use_container_width=True,hide_index=True)
        fig=px.bar(cpm,x="Aktivite",y="Süre (gün)",color="Kritik mi?",color_discrete_map={"EVET":RED,"Hayır":TEAL})
        st.plotly_chart(chart_layout(fig,400),use_container_width=True)
        dated=schedule.dropna(subset=["Başlangıç","Bitiş"]).copy()
        if not dated.empty:
            fig=px.timeline(dated,x_start="Başlangıç",x_end="Bitiş",y="Aktivite",color="Tamamlanma (%)",color_continuous_scale="Blues")
            fig.update_yaxes(autorange="reversed")
            st.plotly_chart(chart_layout(fig,450),use_container_width=True)

# ------------------------------ EVM ------------------------------
elif page == "EVM Performans":
    st.title("Earned Value Management (EVM)")
    st.caption(f"Raporlama tarihi: {report_date:%d.%m.%Y} | BAC = onaylı bütçe toplamı varsayımı")
    cols=st.columns(5)
    for col,label,key in zip(cols,["BAC","PV","EV","AC","EAC"],["BAC","PV","EV","AC","EAC"]):
        col.metric(label,money(evm[key],currency) if pd.notna(evm[key]) else "—")
    c1,c2,c3,c4=st.columns(4)
    c1.metric("CPI",f'{evm["CPI"]:.3f}' if pd.notna(evm["CPI"]) else "—")
    c2.metric("SPI",f'{evm["SPI"]:.3f}' if pd.notna(evm["SPI"]) else "—")
    c3.metric("CV",money(evm["CV"],currency))
    c4.metric("SV",money(evm["SV"],currency))
    st.write("**Yorumlama:** CPI > 1 maliyet verimliliğinin, SPI > 1 takvim performansının planın üzerinde olduğunu gösterir. CPI/SPI hesapları için iş programı ve bütçe WBS kodlarını eşleştirin.")
    st.caption("Bu uygulamadaki PV, aktivite başlangıç-bitiş tarihleri arasında doğrusal zaman ilerlemesi varsayar; kaynak yüklemeli detaylı baz çizgi / onaylı revizyon takvimi bulunmadıkça yaklaşık yönetim göstergesidir.")
    if schedule.empty: st.warning("EVM için önce WBS bazlı iş programı ve tamamlanma yüzdelerini girin.")
    elif plan["WBS"].fillna("").astype(str).str.strip().eq("").any(): st.warning("Bütçede WBS boş satırlar var; EVM sonuçları eksik kalabilir.")
    st.dataframe(pd.DataFrame([{"Gösterge":k,"Değer":v} for k,v in evm.items()]),use_container_width=True,hide_index=True)

# ------------------------------ SATIN ALMA ------------------------------
elif page == "Satın Alma ve Tedarik":
    st.title("Satın Alma ve Tedarik Yönetimi")
    edited=st.data_editor(DATA["purchase"],num_rows="dynamic",use_container_width=True,hide_index=True,
        column_config={"_id":None,"Sipariş Tarihi":st.column_config.DateColumn(),"Termin Tarihi":st.column_config.DateColumn(),
            "Sipariş Miktarı":st.column_config.NumberColumn(min_value=0,format="%.3f"),"Teslim Miktarı":st.column_config.NumberColumn(min_value=0,format="%.3f"),
            "Birim Fiyat":st.column_config.NumberColumn(min_value=0,format="%.4f"),"Sipariş Tutarı":st.column_config.NumberColumn(disabled=True,format="%.2f"),
            "Durum":st.column_config.SelectboxColumn(options=["Talep","Teklif Bekleniyor","Sipariş Verildi","Kısmi Teslimat","Teslim Edildi","İptal"] )},key="purchase_editor")
    if st.button("Satın alma kayıtlarını kaydet",type="primary"): save_from_editor("purchase",edited,PURCHASE_COLS,"purchase")
    pur=DATA["purchase"].copy()
    if not pur.empty:
        pur["Sipariş Tutarı"]=pd.to_numeric(pur["Sipariş Tutarı"],errors="coerce").fillna(0)
        pur["Teslim Oranı (%)"]=np.where(pd.to_numeric(pur["Sipariş Miktarı"],errors="coerce").fillna(0)>0,pd.to_numeric(pur["Teslim Miktarı"],errors="coerce").fillna(0)/pd.to_numeric(pur["Sipariş Miktarı"],errors="coerce").fillna(0)*100,0)
        a,b,c=st.columns(3); a.metric("Toplam Sipariş",money(pur["Sipariş Tutarı"].sum(),currency)); b.metric("Faturalanan",money(pd.to_numeric(pur["Faturalanan Tutar"],errors="coerce").fillna(0).sum(),currency)); c.metric("Ödenen",money(pd.to_numeric(pur["Ödenen Tutar"],errors="coerce").fillna(0).sum(),currency))
        st.dataframe(pur.drop(columns=["_id"],errors="ignore"),use_container_width=True,hide_index=True)

# ------------------------------ PUANTAJ ------------------------------
elif page == "İşçilik ve Puantaj":
    st.title("İşçilik ve Puantaj Yönetimi")
    st.info("Toplam İşçilik = Çalışan Sayısı × Gün × Kişi Günlük Ücret. Puantaj kaydı gerçekleşen maliyete otomatik aktarılmaz; çift kayıt oluşmasını önlemek için aktarım kullanıcı kontrolündedir.")
    edited=st.data_editor(DATA["labor"],num_rows="dynamic",use_container_width=True,hide_index=True,
        column_config={"_id":None,"Tarih":st.column_config.DateColumn(),"Çalışan Sayısı":st.column_config.NumberColumn(min_value=0),"Gün":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Kişi Günlük Ücret":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Toplam İşçilik":st.column_config.NumberColumn(disabled=True,format="%.2f")},key="labor_editor")
    if st.button("Puantaj kayıtlarını kaydet",type="primary"): save_from_editor("labor",edited,LABOR_COLS,"labor")
    labor=DATA["labor"]
    if not labor.empty:
        total=pd.to_numeric(labor["Toplam İşçilik"],errors="coerce").fillna(0).sum()
        man_days=(pd.to_numeric(labor["Çalışan Sayısı"],errors="coerce").fillna(0)*pd.to_numeric(labor["Gün"],errors="coerce").fillna(0)).sum()
        a,b=st.columns(2); a.metric("Toplam İşçilik",money(total,currency)); b.metric("Toplam Adam-Gün",f"{man_days:,.1f}")
        st.caption("Puantajı gerçekleşen maliyete aktarmak için aynı tutarı ayrıca girin; otomatik aktarım yapılmaz.")

# ------------------------------ DEPO / STOK ------------------------------
elif page == "Depo / Stok":
    st.title("Depo ve Stok Hareketleri")
    st.caption("Giriş ve sayım artısı stoğu artırır; çıkış ve sayım eksisi azaltır. Miktarı pozitif girin; hareket türü işareti belirler. Eski 'Düzeltme' hareketleri artı düzeltme olarak yorumlanır.")
    edited=st.data_editor(DATA["inventory"],num_rows="dynamic",use_container_width=True,hide_index=True,
        column_config={"_id":None,"Tarih":st.column_config.DateColumn(),"Hareket Türü":st.column_config.SelectboxColumn(options=["Giriş","Çıkış","Sayım Artısı","Sayım Eksisi","Düzeltme"]),"Miktar":st.column_config.NumberColumn(min_value=0,format="%.3f"),"Birim Fiyat":st.column_config.NumberColumn(min_value=0,format="%.4f"),"Tutar":st.column_config.NumberColumn(disabled=True,format="%.2f")},key="inventory_editor")
    if st.button("Stok hareketlerini kaydet",type="primary"): save_from_editor("inventory",edited,INVENTORY_COLS,"inventory")
    inv=DATA["inventory"].copy()
    if not inv.empty:
        inv["Miktar"] = pd.to_numeric(inv["Miktar"],errors="coerce").fillna(0)
        movement = inv["Hareket Türü"].astype(str)
        sign = np.where(movement.isin(["Çıkış", "Sayım Eksisi"]), -1, 1)
        inv["Net Miktar"] = sign * inv["Miktar"]
        inv["Net Değer"] = sign * pd.to_numeric(inv["Tutar"],errors="coerce").fillna(0)
        balances=inv.groupby(["Depo","Malzeme Kodu","Malzeme","Birim"],dropna=False).agg(Stok=("Net Miktar","sum"),Stok_Değeri=("Net Değer","sum")).reset_index()
        st.subheader("Güncel Stok Bakiyesi"); st.dataframe(balances,use_container_width=True,hide_index=True)
        neg=balances[balances["Stok"]<0]
        if not neg.empty: st.error(f"Negatif bakiye bulunan {len(neg)} malzeme var. Depo hareketlerini kontrol edin.")
        st.subheader("Hareket Dökümü"); st.dataframe(inv.drop(columns=["_id"],errors="ignore"),use_container_width=True,hide_index=True)

# ------------------------------ HAKEDİŞ ------------------------------
elif page == "Hakediş Takibi":
    st.title("Taşeron Hakediş Takibi")
    st.caption("Net Hakediş = Brüt Tutar − Kesinti. Ödenecek Tutar = Net Hakediş + KDV. KDV oranı 0,20 gibi ondalık girilir.")
    edited=st.data_editor(DATA["hakediş"],num_rows="dynamic",use_container_width=True,hide_index=True,
        column_config={"_id":None,"Tarih":st.column_config.DateColumn(),"Brüt Tutar":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Kesinti":st.column_config.NumberColumn(min_value=0,format="%.2f"),"KDV Oranı":st.column_config.NumberColumn(min_value=0,max_value=1,format="%.2f"),"Net Hakediş":st.column_config.NumberColumn(disabled=True,format="%.2f"),"KDV Tutarı":st.column_config.NumberColumn(disabled=True,format="%.2f"),"Ödenecek Tutar":st.column_config.NumberColumn(disabled=True,format="%.2f"),"Ödenen":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Durum":st.column_config.SelectboxColumn(options=["Hazırlanıyor","Kontrolde","Onaylandı","Kısmi Ödendi","Ödendi","İtirazlı"] )},key="hak_editor")
    if st.button("Hakediş kayıtlarını kaydet",type="primary"): save_from_editor("hakediş",edited,HAKEDIS_COLS,"hakediş")
    h=DATA["hakediş"].copy()
    if not h.empty:
        a,b,c=st.columns(3); a.metric("Toplam Net Hakediş",money(pd.to_numeric(h["Net Hakediş"],errors="coerce").fillna(0).sum(),currency)); b.metric("Toplam Ödenecek",money(pd.to_numeric(h["Ödenecek Tutar"],errors="coerce").fillna(0).sum(),currency)); c.metric("Toplam Ödenen",money(pd.to_numeric(h["Ödenen"],errors="coerce").fillna(0).sum(),currency))
        st.dataframe(h.drop(columns=["_id"],errors="ignore"),use_container_width=True,hide_index=True)

# ------------------------------ NAKİT AKIŞI / S-CURVE ------------------------------
elif page == "Nakit Akışı / S-Curve":
    st.title("Nakit Akışı ve S-Curve")
    if not DATA.get("legacy_cash", pd.DataFrame()).empty:
        with st.expander("V11.0 eski dönemsel nakit akışı arşivi", expanded=False):
            st.caption("Bu tablo V11.0 biçiminden korunarak kopyalanmıştır; yeni tahsilat/ödeme hareketleriyle otomatik birleştirilmez.")
            st.dataframe(DATA["legacy_cash"].drop(columns=["_id"],errors="ignore"),use_container_width=True,hide_index=True)
    edited=st.data_editor(DATA["cash"],num_rows="dynamic",use_container_width=True,hide_index=True,
        column_config={"_id":None,"Tarih":st.column_config.DateColumn(),"Vade Tarihi":st.column_config.DateColumn(),"Tür":st.column_config.SelectboxColumn(options=["Tahsilat","Ödeme"]),"Tutar":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Durum":st.column_config.SelectboxColumn(options=["Planlandı","Gerçekleşti","Gecikti"] )},key="cash_editor")
    if st.button("Nakit akışı kayıtlarını kaydet",type="primary"):
        frame=ensure_frame(edited,CASH_COLS); frame["Tutar"]=pd.to_numeric(frame["Tutar"],errors="coerce").fillna(0); save_from_editor("cash",frame,CASH_COLS)
    cash=DATA["cash"].copy()
    if not cash.empty:
        cash["Tutar"]=pd.to_numeric(cash["Tutar"],errors="coerce").fillna(0)
        realized=cash[cash["Durum"].astype(str)=="Gerçekleşti"]
        inflow=realized.loc[realized["Tür"]=="Tahsilat","Tutar"].sum(); outflow=realized.loc[realized["Tür"]=="Ödeme","Tutar"].sum()
        a,b,c=st.columns(3); a.metric("Gerçekleşen Tahsilat",money(inflow,currency)); b.metric("Gerçekleşen Ödeme",money(outflow,currency)); c.metric("Net Nakit",money(inflow-outflow,currency))
        cash["Tarih_dt"]=pd.to_datetime(cash["Tarih"],errors="coerce",format="mixed",dayfirst=True)
        valid=cash.dropna(subset=["Tarih_dt"]).sort_values("Tarih_dt").copy()
        if not valid.empty:
            valid["Net Akış"]=np.where(valid["Tür"]=="Tahsilat",valid["Tutar"],-valid["Tutar"])
            daily=valid.groupby("Tarih_dt")["Net Akış"].sum().reset_index(); daily["Kümülatif Net"] = daily["Net Akış"].cumsum()
            fig=go.Figure(); fig.add_trace(go.Scatter(x=daily["Tarih_dt"],y=daily["Kümülatif Net"],mode="lines+markers",name="Kümülatif Net Nakit",line=dict(color=TEAL,width=3)))
            st.plotly_chart(chart_layout(fig,400),use_container_width=True)
            valid["Kümülatif Tutar"]=valid.groupby("Tür")["Tutar"].cumsum()
            fig=px.line(valid,x="Tarih_dt",y="Kümülatif Tutar",color="Tür",markers=True)
            st.plotly_chart(chart_layout(fig,400),use_container_width=True)
    else:
        st.info("Nakit akışı kaydı bulunmuyor.")

    st.subheader("Maliyet S-Curve — Planlanan / Gerçekleşen")
    st.caption("Planlanan eğri, ilgili WBS bütçesinin aktivite başlangıç-bitiş aylarına eşit dağıtılmasıyla oluşturulan yaklaşık baz çizgidir; detaylı aylık nakit yüklemesi varsa onunla değiştirilmelidir.")
    if not plan.empty and not actual.empty and not schedule.empty:
        p_curve = plan.copy()
        a_curve = actual.copy()
        p_curve["WBS"] = p_curve["WBS"].fillna("").astype(str).str.strip()
        a_curve["Tarih_dt"] = pd.to_datetime(a_curve["Tarih"], errors="coerce", format="mixed", dayfirst=True)
        s_curve = schedule.copy()
        s_curve["WBS"] = s_curve["WBS"].fillna("").astype(str).str.strip()
        s_curve["Başlangıç"] = pd.to_datetime(s_curve["Başlangıç"], errors="coerce", format="mixed", dayfirst=True)
        s_curve["Bitiş"] = pd.to_datetime(s_curve["Bitiş"], errors="coerce", format="mixed", dayfirst=True)
        s_curve = s_curve.dropna(subset=["Başlangıç","Bitiş"])
        if not s_curve.empty and a_curve["Tarih_dt"].notna().any():
            start_curve = min(s_curve["Başlangıç"].min(), a_curve["Tarih_dt"].min())
            end_curve = max(s_curve["Bitiş"].max(), a_curve["Tarih_dt"].max())
            if start_curve == end_curve: end_curve = end_curve + pd.Timedelta(days=1)
            months = pd.date_range(start=start_curve.to_period("M").start_time, end=end_curve.to_period("M").start_time, freq="MS")
            plan_monthly = pd.Series(0.0,index=months)
            for wbs, group in s_curve.groupby("WBS"):
                if not wbs: continue
                amount = p_curve.loc[p_curve["WBS"]==wbs,"Planlanan Tutar"].sum()
                first,last = group["Başlangıç"].min(),group["Bitiş"].max()
                if pd.isna(first) or pd.isna(last): continue
                active_months = months[(months >= first.to_period("M").start_time) & (months <= last.to_period("M").start_time)]
                if len(active_months):
                    for month in active_months: plan_monthly.loc[month] += amount/len(active_months)
            actual_monthly = a_curve.dropna(subset=["Tarih_dt"]).groupby(a_curve.dropna(subset=["Tarih_dt"])["Tarih_dt"].dt.to_period("M"))["Gerçekleşen Tutar"].sum()
            curve = pd.DataFrame({"Ay":months,"Planlanan Kümülatif":plan_monthly.reindex(months,fill_value=0).cumsum().values,
                                  "Gerçekleşen Kümülatif":[actual_monthly.get(m.to_period("M"),0.0) for m in months]})
            curve["Gerçekleşen Kümülatif"] = curve["Gerçekleşen Kümülatif"].cumsum()
            fig=go.Figure()
            fig.add_trace(go.Scatter(x=curve["Ay"],y=curve["Planlanan Kümülatif"],mode="lines+markers",name="Planlanan",line=dict(color=NAVY,width=3)))
            fig.add_trace(go.Scatter(x=curve["Ay"],y=curve["Gerçekleşen Kümülatif"],mode="lines+markers",name="Gerçekleşen",line=dict(color=RED,width=3)))
            st.plotly_chart(chart_layout(fig,420),use_container_width=True)
        else:
            st.info("S-Curve için program tarihleri ve tarih girilmiş gerçekleşen maliyet kaydı gerekir.")
    else:
        st.info("Maliyet S-Curve için bütçe, WBS bazlı iş programı ve tarihli gerçekleşen maliyet verisi gerekir.")

# ------------------------------ RİSK YÖNETİMİ ------------------------------
elif page == "Risk Yönetimi":
    st.title("Risk Yönetimi ve Önceliklendirme")
    edited=st.data_editor(risk.drop(columns=["Risk Skoru"],errors="ignore"),num_rows="dynamic",use_container_width=True,hide_index=True,
        column_config={"_id":None,"Olasılık (1-5)":st.column_config.NumberColumn(min_value=1,max_value=5,step=1),"Etki (1-5)":st.column_config.NumberColumn(min_value=1,max_value=5,step=1),"Tahmini Maliyet Etkisi":st.column_config.NumberColumn(min_value=0,format="%.2f"),"Durum":st.column_config.SelectboxColumn(options=["Açık","İzleniyor","Kapalı"]),"Hedef Tarih":st.column_config.DateColumn()},key="risk_editor")
    if st.button("Riskleri güvenli kaydet",type="primary"): save_from_editor("risk",edited,RISK_COLS,"risk")
    if not risk.empty:
        st.subheader("Risk Matrisi")
        fig=px.scatter(risk,x="Olasılık (1-5)",y="Etki (1-5)",size="Tahmini Maliyet Etkisi",color="Risk Skoru",hover_name="Risk Tanımı",text="Risk ID",range_x=[0.5,5.5],range_y=[0.5,5.5],color_continuous_scale="RdYlGn_r")
        st.plotly_chart(chart_layout(fig,480),use_container_width=True)
        exposure=(pd.to_numeric(risk["Tahmini Maliyet Etkisi"],errors="coerce").fillna(0)*pd.to_numeric(risk["Olasılık (1-5)"],errors="coerce").fillna(0)/5).sum()
        st.metric("Olasılık Ağırlıklı Risk Maliyeti",money(exposure,currency))
        st.dataframe(risk.drop(columns=["_id"],errors="ignore"),use_container_width=True,hide_index=True)
    else: st.info("Henüz risk kaydı yok.")

# ------------------------------ VERİ KONTROLÜ ------------------------------
elif page == "Veri Kontrolü":
    st.title("Veri Kalitesi ve Tutarlılık Kontrolleri")
    checks=calculate_quality_checks(plan,actual,schedule,risk,DATA["library"])
    st.dataframe(checks,use_container_width=True,hide_index=True)
    total=int(checks["Uyarı Sayısı"].sum()) if not checks.empty else 0
    if total==0: st.success("Tanımlı temel veri kontrollerinde uyarı bulunmadı.")
    else: st.warning(f"{total} adet veri kontrolü uyarısı bulundu.")
    st.caption("Bu kontroller veri giriş tutarlılığını değerlendirir; statik proje, metraj, piyasa fiyatı veya saha imalatının teknik doğruluğunu tasdik etmez.")
    st.subheader("CPM Bağlantı Kontrolleri")
    for msg in cpm_warnings: st.warning(msg)
    if not cpm.empty: st.dataframe(cpm[["WBS","Aktivite","Öncül WBS","Kritik mi?","Bolluk (gün)"]],use_container_width=True,hide_index=True)

# ------------------------------ YEDEKLEME / DENETİM ------------------------------
elif page == "Yedekleme / Denetim":
    st.title("Yedekleme, Snapshot ve Denetim İzleri")
    snapshot_sets={ds:DATA[ds] for ds in DATA if ds!="library"}
    if st.button("Bu projenin anlık yedeğini (snapshot) oluştur",type="primary"):
        snap_id=make_snapshot(project_key,snapshot_sets)
        audit_log(project_key,"SNAPSHOT","all",sum(len(x) for x in snapshot_sets.values()),f"Snapshot ID: {snap_id}")
        st.success(f"Snapshot oluşturuldu: {snap_id}")
    export_json={"app":APP_NAME,"version":APP_VERSION,"project":st.session_state.current_project,"created_at":now_tr().isoformat(),"data":{k:v.to_dict(orient="records") for k,v in snapshot_sets.items()},"library":DATA["library"].to_dict(orient="records")}
    st.download_button("Tam proje yedeğini JSON indir",json.dumps(json_safe_value(export_json),ensure_ascii=False,allow_nan=False,indent=2).encode("utf-8"),file_name=f"Emir_Asel_{re.sub(r'[^A-Za-z0-9_-]','_',st.session_state.current_project)}_backup_{now_tr():%Y%m%d_%H%M}.json",mime="application/json")
    restore_file = st.file_uploader("Daha önce alınmış V12 JSON yedeğini seç", type=["json"], key="restore_backup")
    if restore_file and st.button("Seçili yedeği bu projeye geri yükle", type="secondary"):
        try:
            backup = json.loads(restore_file.getvalue().decode("utf-8"))
            backup_data = backup.get("data", {})
            if not isinstance(backup_data, dict): raise ValueError("Yedek dosyasında data alanı bulunamadı.")
            # Önce tüm setleri doğrula; sonra ayrı dataset transaction'larıyla yaz.
            restored = {}
            for ds, cols in DATASETS.items():
                if ds in ("project", "library"): continue
                raw_rows = backup_data.get(ds, [])
                if not isinstance(raw_rows, list): raise ValueError(f"{ds} veri seti liste biçiminde değil.")
                restored[ds] = normalize_dataset_dates(ensure_frame(pd.DataFrame(raw_rows), cols), ds)
            for ds, frame in restored.items(): save_dataset(project_key, ds, frame, DATASETS[ds])
            for ds, frame in restored.items(): st.session_state[f"data_{ds}"] = frame
            # Merkezi fiyat kütüphanesi geri yüklemede silinmez; yedekteki yeni kayıtlar eklenir.
            backup_library = backup.get("library", [])
            if isinstance(backup_library, list) and backup_library:
                library_add = ensure_frame(pd.DataFrame(backup_library), LIB_COLS)
                current_library = load_dataset("__GLOBAL__", "library", LIB_COLS)
                existing_ids = set(current_library["_id"].astype(str))
                library_add = library_add[~library_add["_id"].astype(str).isin(existing_ids)]
                merged_library = pd.concat([current_library, library_add], ignore_index=True)
                save_dataset("__GLOBAL__", "library", merged_library, LIB_COLS)
                st.session_state["data_library"] = ensure_frame(merged_library, LIB_COLS)
            audit_log(project_key,"YEDEKTEN GERİ YÜKLE", "all", sum(len(v) for v in restored.values()), restore_file.name)
            st.success("Yedek verileri geri yüklendi."); st.rerun()
        except Exception as exc: st.error(f"Yedek geri yüklenemedi: {exc}")
    st.subheader("Denetim Kayıtları")
    logs=load_dataset(project_key,"audit",["_id","Tarih","Kullanıcı","İşlem","Veri Seti","Kayıt Sayısı","Detay"])
    if not logs.empty: st.dataframe(logs.drop(columns=["_id"],errors="ignore").sort_values("Tarih",ascending=False),use_container_width=True,hide_index=True)
    else: st.info("Henüz denetim kaydı bulunmuyor.")
    st.subheader("Snapshot Kayıtları")
    snaps=load_dataset(project_key,"snapshots",["_id","Tarih","Veri"])
    if not snaps.empty:
        summary=pd.DataFrame([{"Tarih":r.get("Tarih"),"Snapshot ID":r.get("_id"),"Veri Seti Sayısı":len(r.get("Veri",{}))} for _,r in snaps.iterrows()])
        st.dataframe(summary,use_container_width=True,hide_index=True)

# ------------------------------ EXCEL RAPORLARI ------------------------------
elif page == "Excel Raporları":
    st.title("Excel Raporları ve Dışa Aktarma")
    report_tables={
        "Proje_Ozeti":pd.DataFrame([["Proje",st.session_state.current_project],["Sürüm",APP_VERSION],["Rapor Tarihi",report_date.isoformat()],["Para Birimi",currency],["Toplam Bütçe",budget],["Gerçekleşen Maliyet",actual_cost],["Fiziksel İlerleme (%)",progress],["BAC",evm["BAC"]],["PV",evm["PV"]],["EV",evm["EV"]],["AC",evm["AC"]],["CPI",evm["CPI"]],["SPI",evm["SPI"]],["EAC",evm["EAC"]]],columns=["Gösterge","Değer"]),
        "Butce_Metraj":plan.drop(columns=["_id"],errors="ignore"),"Gerceklesen":actual.drop(columns=["_id"],errors="ignore"),
        "Kategori_Sapma":cost_summary,"Is_Programi":schedule.drop(columns=["_id"],errors="ignore"),"CPM":cpm.drop(columns=["_id"],errors="ignore"),
        "Riskler":risk.drop(columns=["_id"],errors="ignore"),"Birim_Fiyat":DATA["library"].drop(columns=["_id"],errors="ignore"),
        "Satin_Alma":DATA["purchase"].drop(columns=["_id"],errors="ignore"),"Puantaj":DATA["labor"].drop(columns=["_id"],errors="ignore"),
        "Depo_Stok":DATA["inventory"].drop(columns=["_id"],errors="ignore"),"Hakedis":DATA["hakediş"].drop(columns=["_id"],errors="ignore"),
        "Nakit_Akisi":DATA["cash"].drop(columns=["_id"],errors="ignore"),"V11_Nakit_Arsiv":DATA["legacy_cash"].drop(columns=["_id"],errors="ignore"),"Veri_Kontrol":calculate_quality_checks(plan,actual,schedule,risk,DATA["library"])}
    xlsx=excel_export(report_tables)
    st.download_button("📥 Profesyonel Excel raporunu indir",xlsx,file_name=f"Emir_Asel_V12_{re.sub(r'[^A-Za-z0-9_-]','_',st.session_state.current_project)}_{now_tr():%Y%m%d_%H%M}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
    st.subheader("Rapor Ön İzleme — Kategori Sapmaları")
    st.dataframe(cost_summary,use_container_width=True,hide_index=True)

st.divider()
st.caption(f"{APP_NAME} | Proje Kontrol V{APP_VERSION} | {DB_KIND.upper()} veri tabanı | Karar destek amaçlıdır; resmi mühendislik/muhasebe onayının yerine geçmez.")
