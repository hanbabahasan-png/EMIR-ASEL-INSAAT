# ============================================================
# EMİR ASEL İNŞAAT | PROJE KONTROL SİSTEMİ V7 — PREMIUM
# ============================================================
# Kurulum:   pip install streamlit pandas openpyxl plotly numpy
# Çalıştır:  streamlit run app_v7.py
# ============================================================

from io import BytesIO
from datetime import datetime
import re
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment

# ------------------------------------------------------------
# SAYFA AYARI
# ------------------------------------------------------------
st.set_page_config(
    page_title="Emir Asel | Proje Kontrol V7",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_VERSION = "7.0"
VAT_DEFAULT = 0.20
RISK_DEFAULT = 0.10

# ------------------------------------------------------------
# PREMIUM CHART THEME
# ------------------------------------------------------------
BRAND = {
    "primary": "#0A192F",
    "accent":  "#E63946",
    "success": "#2A9D8F",
    "warning": "#F4A261",
    "info":    "#457B9D",
    "gold":    "#E9C46A",
    "purple":  "#7209B7",
    "gray":    "#6C757D",
    "grid":    "#E9ECEF",
}

CHART_SEQUENCE = [
    BRAND["primary"], BRAND["accent"], BRAND["success"],
    BRAND["warning"], BRAND["info"], BRAND["gold"], BRAND["purple"],
]


def apply_theme(fig, height=400, show_legend=True, title=None):
    fig.update_layout(
        height=height,
        template="plotly_white",
        font=dict(family="Inter, -apple-system, Segoe UI, sans-serif",
                  size=13, color="#1A1A1A"),
        title=dict(
            text=title or (fig.layout.title.text if fig.layout.title else ""),
            font=dict(size=16, color=BRAND["primary"], family="Inter"),
            x=0.01, xanchor="left", pad=dict(b=12),
        ),
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02,
            xanchor="right", x=1,
            bgcolor="rgba(255,255,255,0.9)",
            bordercolor=BRAND["grid"], borderwidth=1,
            visible=show_legend,
        ),
        margin=dict(l=16, r=16, t=60, b=16),
        plot_bgcolor="white",
        paper_bgcolor="white",
        hoverlabel=dict(
            bgcolor=BRAND["primary"], font_size=12,
            font_family="Inter", font_color="white",
        ),
    )
    fig.update_xaxes(
        showgrid=False, showline=True, linecolor=BRAND["grid"],
        tickfont=dict(size=11, color="#4A4A4A"),
    )
    fig.update_yaxes(
        showgrid=True, gridcolor=BRAND["grid"], gridwidth=1,
        zeroline=False, showline=False,
        tickfont=dict(size=11, color="#4A4A4A"),
    )
    return fig


def fmt_money(v):
    if v is None or pd.isna(v):
        return "—"
    if abs(v) >= 1_000_000:
        return f"₺{v/1_000_000:.2f}M"
    if abs(v) >= 1_000:
        return f"₺{v/1_000:.0f}K"
    return f"₺{v:.0f}"


# ------------------------------------------------------------
# SABİTLER / ALIASES
# ------------------------------------------------------------
SHEET_ALIASES = {
    "plan":       ["Planlanan Maliyet", "01_IS_KALEMLERI", "İş Kalemleri"],
    "actual":     ["Gerçek Maliyet", "03_GERCEK_MALIYET", "Gerçekleşen Maliyet"],
    "schedule":   ["İş Programı", "05_IS_PROGRAMI", "Is Programi"],
    "finance":    ["Finans", "06_FINANS", "Nakit Akışı"],
    "risk":       ["Risk", "07_RISK_KAPSAM", "Risk Kapsam"],
    "procurement":["Satın Alma", "04_SATIN_ALMA", "Satinalma"],
}

ALIASES = {
    "wbs":             ["WBS", "WBS Kodu", "Kod", "İş Kalemi Kodu"],
    "item":            ["Kalem", "İş Kalemi", "İş Kalemi / İmalat Tanımı", "Açıklama"],
    "category":        ["Kategori", "Disiplin", "Disiplin / Kategori", "Ana Kalem"],
    "unit":            ["Birim"],
    "qty":             ["Miktar", "Baz Miktar", "Metraj", "Planlanan Miktar"],
    "status":          ["Statü", "Miktar Statüsü", "Durum"],
    "unit_price":      ["SEÇİLEN BF", "Bütçe BF (KDV hariç)", "Birim Fiyat"],
    "budget":          ["PLANLANAN TUTAR", "Bütçe Tutarı", "Planlanan Tutar", "BAC"],
    "actual_qty":      ["Fiili Miktar", "Gerçek Miktar", "Gerçekleşen Miktar"],
    "actual_price":    ["Gerçek Birim Fiyat", "Fiili Birim Fiyat"],
    "actual_total":    ["Toplam", "Gerçekleşen Tutar", "Tutar", "Ara Toplam"],
    "progress":        ["Gerçekleşen İlerleme", "Fiili İlerleme", "Gerçekleşme %", "İlerleme %"],
    "planned_progress":["Planlanan İlerleme", "Planlanan İlerleme %", "Plan %"],
    "weight":          ["Ağırlık", "İş Kalemi Ağırlığı", "Weight"],
    "date":            ["Tarih", "İşlem Tarihi", "Fatura Tarihi"],
    "amount":          ["Tutar", "Tutar (TL)", "Gerçekleşen Tutar"],
    "direction":       ["Hareket Türü", "Tür", "İşlem Türü"],
    "predecessor":     ["Öncül Faaliyet", "Öncül Faaliyetler", "Öncel Faaliyetler"],
    "duration":        ["Süre", "Süre (gün)", "Süre (iş günü)"],
    "activity":        ["Aktivite Kodu", "Faaliyet Kodu", "Aktivite ID"],
    "probability":     ["Olasılık", "Olasilik", "Probability"],
    "impact":          ["Etki", "Etki Puanı", "Impact"],
    "risk_level":      ["Risk Seviyesi", "Önem", "Risk Düzeyi"],
}


# ------------------------------------------------------------
# YARDIMCILAR
# ------------------------------------------------------------
def norm(value):
    value = str(value).strip().lower()
    for old, new in {"ı":"i","İ":"i","ş":"s","ğ":"g","ü":"u","ö":"o","ç":"c"}.items():
        value = value.replace(old, new)
    return re.sub(r"[^a-z0-9]", "", value)


def find_col(df, key):
    candidates = ALIASES.get(key, [key])
    lookup = {norm(c): c for c in df.columns}
    for candidate in candidates:
        found = lookup.get(norm(candidate))
        if found is not None:
            return found
    return None


def numeric(series):
    if series is None or len(series) == 0:
        return pd.Series(dtype=float)
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    s = series.astype(str).str.strip()
    s = s.replace({"": np.nan, "None": np.nan, "nan": np.nan,
                   "NaN": np.nan, "-": np.nan})
    direct = pd.to_numeric(s, errors="coerce")
    failed = direct.isna() & s.notna()
    if failed.any():
        s_try = (s[failed]
                 .str.replace(".", "", regex=False)
                 .str.replace(",", ".", regex=False))
        direct.loc[failed] = pd.to_numeric(s_try, errors="coerce")
    return direct


def get_sheet(raw_sheets, logical_name):
    for wanted in SHEET_ALIASES.get(logical_name, []):
        for actual_name, df in raw_sheets.items():
            if norm(actual_name) == norm(wanted):
                return df.copy()
    return pd.DataFrame()


def money(value, short=False):
    if value is None or pd.isna(value):
        return "—"
    if short:
        return (f"₺{value/1_000_000:,.2f}M"
                if abs(value) >= 1_000_000 else f"₺{value:,.0f}")
    return f"₺{value:,.2f}"


# ------------------------------------------------------------
# EXCEL OKUMA (CACHE)
# ------------------------------------------------------------
@st.cache_data(show_spinner="Excel dosyası okunuyor...")
def load_excel(file_bytes: bytes, file_name: str):
    excel = pd.ExcelFile(BytesIO(file_bytes))
    raw_sheets = {
        name: pd.read_excel(BytesIO(file_bytes), sheet_name=name)
        for name in excel.sheet_names
    }
    return raw_sheets, excel.sheet_names


# ------------------------------------------------------------
# VERİ HAZIRLAMA
# ------------------------------------------------------------
PLAN_REQUIRED = {
    "wbs": "", "item": "", "category": "Sınıflandırılmamış",
    "unit": "", "status": "Bilinmiyor",
    "qty": np.nan, "unit_price": np.nan, "budget": np.nan,
}


def prepare_plan(df):
    if df.empty:
        return pd.DataFrame()
    out = df.copy()
    for key in ["wbs","item","category","unit","qty","status",
                "unit_price","budget"]:
        col = find_col(out, key)
        if col and col != key:
            out[key] = out[col]
    for col in ["qty","unit_price","budget"]:
        if col in out:
            out[col] = numeric(out[col])
    for col, default in PLAN_REQUIRED.items():
        if col not in out:
            out[col] = default
    calc = out["budget"].isna() & out["qty"].notna() & out["unit_price"].notna()
    if calc.any():
        out.loc[calc, "budget"] = (out.loc[calc, "qty"]
                                   * out.loc[calc, "unit_price"])
    out["wbs"] = out["wbs"].fillna("").astype(str).str.strip()
    out["item"] = out["item"].fillna("").astype(str)
    out["category"] = out["category"].fillna("Sınıflandırılmamış").astype(str)
    out["status"] = out["status"].fillna("Bilinmiyor").astype(str)
    out["unit"] = out["unit"].fillna("").astype(str)
    out = out[(out["wbs"] != "") & (out["item"].str.strip() != "")].copy()
    out["wbs_duplicate"] = out["wbs"].duplicated(keep=False)
    return out


def prepare_actual(df):
    if df.empty:
        return pd.DataFrame()
    out = df.copy()
    for key in ["wbs","item","actual_qty","actual_price","actual_total","date"]:
        col = find_col(out, key)
        if col and col != key:
            out[key] = out[col]
    for col in ["actual_qty","actual_price","actual_total"]:
        if col in out:
            out[col] = numeric(out[col])
    if "actual_total" not in out:
        if "actual_qty" in out and "actual_price" in out:
            out["actual_total"] = out["actual_qty"] * out["actual_price"]
        else:
            out["actual_total"] = np.nan
    for col in ["wbs","item"]:
        if col not in out:
            out[col] = ""
    out["wbs"] = out["wbs"].fillna("").astype(str).str.strip()
    out["item"] = out["item"].fillna("").astype(str)
    out["actual_total"] = out["actual_total"].fillna(0)
    out["_date"] = (pd.to_datetime(out["date"], errors="coerce", dayfirst=True)
                    if "date" in out else pd.NaT)
    return out


def prepare_schedule(df):
    if df.empty:
        return pd.DataFrame()
    out = df.copy()
    for key in ["wbs","item","progress","planned_progress","weight",
                "activity","predecessor","duration"]:
        col = find_col(out, key)
        if col and col != key:
            out[key] = out[col]
    for col in ["progress","planned_progress","weight","duration"]:
        if col in out:
            out[col] = numeric(out[col])
    for col in ["progress","planned_progress"]:
        if col in out:
            mx = out[col].max()
            if pd.notna(mx) and mx <= 1.0:
                out[col] = out[col] * 100
            out[col] = out[col].where(out[col].between(0, 100), np.nan)
    if "wbs" not in out: out["wbs"] = ""
    if "item" not in out: out["item"] = ""
    if "activity" not in out: out["activity"] = out["wbs"]
    out["wbs"] = out["wbs"].fillna("").astype(str).str.strip()
    out["item"] = out["item"].fillna("").astype(str)
    out["activity"] = (out["activity"].fillna(out["wbs"])
                       .astype(str).str.strip())
    return out


def prepare_risk(df):
    if df.empty:
        return pd.DataFrame()
    out = df.copy()
    p = find_col(out, "probability")
    i = find_col(out, "impact")
    out["_probability"] = numeric(out[p]) if p else 0.0
    out["_impact"] = numeric(out[i]) if i else 0.0
    out["_probability"] = out["_probability"].clip(0, 5).fillna(0)
    out["_impact"] = out["_impact"].clip(0, 5).fillna(0)
    out["_risk_score"] = out["_probability"] * out["_impact"]

    def level(s):
        if s >= 20: return "ÇOK YÜKSEK"
        if s >= 15: return "YÜKSEK"
        if s >= 8:  return "ORTA"
        if s > 0:   return "DÜŞÜK"
        return "PUANLANMAMIŞ"
    out["_risk_level"] = out["_risk_score"].apply(level)
    return out


def category_group(category, item=""):
    text = norm(f"{category} {item}")
    if any(x in text for x in ["mimarlik","proje gider","yapi denetim",
                               "laboratuvar","genel gider","santiye",
                               "kalip demir iscilik","elektrik su iscilik"]):
        return "GENEL GİDER / İŞÇİLİK"
    if any(x in text for x in ["hafriyat","grobeton","blokaj","dolgu","beton",
                               "demir","kalip","radye","subasman","kolon",
                               "perde","kiris","tabliye","merdiven"]):
        return "KABA İŞLER"
    if any(x in text for x in ["bahce","peyzaj","otopark","drenaj","dis saha",
                               "kanalizasyon","cevre beton"]):
        return "DIŞ İŞLER"
    if any(x in text for x in ["siva","alci","seramik","boya","mekanik",
                               "elektrik","tesisat","cephe","dograma",
                               "yalitim","cati","ince is"]):
        return "İNCE İŞLER / MEP"
    return "SINIFLANDIRILACAK"


def weighted_progress(schedule):
    cols = ["wbs","actual_progress","planned_progress"]
    if schedule.empty or "progress" not in schedule:
        return pd.DataFrame(columns=cols)
    s = schedule[schedule["wbs"] != ""].copy()
    s["progress"] = numeric(s["progress"])
    s["planned_progress"] = (numeric(s["planned_progress"])
                             if "planned_progress" in s else np.nan)
    if "weight" in s and s["weight"].notna().any():
        s["_weight"] = s["weight"].fillna(0)
    else:
        s["_weight"] = 1.0
    s = s[s["progress"].notna()].copy()
    if s.empty:
        return pd.DataFrame(columns=cols)
    s["_wp"] = s["progress"] * s["_weight"]
    s["_wpp"] = s["planned_progress"].fillna(0) * s["_weight"]
    g = s.groupby("wbs").agg(
        wp=("_wp","sum"), wpp=("_wpp","sum"), w=("_weight","sum"),
        avg_prog=("progress","mean"), avg_pp=("planned_progress","mean"),
    ).reset_index()
    g["actual_progress"] = np.where(g["w"] > 0, g["wp"]/g["w"], g["avg_prog"])
    g["planned_progress"] = np.where(g["w"] > 0, g["wpp"]/g["w"], g["avg_pp"])
    return g[cols]


# ------------------------------------------------------------
# CPM
# ------------------------------------------------------------
def calculate_cpm(schedule):
    required = ["activity","duration","predecessor"]
    if schedule.empty or not all(k in schedule for k in required):
        return pd.DataFrame(), "CPM için aktivite, süre ve öncül verisi gerekli."
    df = pd.DataFrame({
        "code": schedule["activity"].astype(str).str.strip(),
        "name": schedule["item"].astype(str),
        "duration": numeric(schedule["duration"]),
        "pred": schedule["predecessor"].fillna("").astype(str),
    })
    df = df[(df["code"] != "") & (df["code"].str.lower() != "nan")].copy()
    df["duration"] = df["duration"].fillna(0).clip(lower=0)
    if df["code"].duplicated().any():
        dup = df[df["code"].duplicated(keep=False)]["code"].unique().tolist()
        return pd.DataFrame(), f"Mükerrer aktivite kodu: {', '.join(dup[:5])}"
    activities = {row["code"]: row for _, row in df.iterrows()}
    pred_map = {c: [] for c in activities}
    succ_map = {c: [] for c in activities}
    for code, row in activities.items():
        preds = [p.strip() for p in re.split(r"[,;|]", row["pred"]) if p.strip()]
        if code in preds:
            return pd.DataFrame(), f"'{code}' kendisine bağlanamaz."
        for p in preds:
            if p not in activities:
                return pd.DataFrame(), f"'{code}' öncülü bulunamadı: {p}"
        pred_map[code] = preds
        for p in preds:
            succ_map[p].append(code)
    indeg = {c: len(pred_map[c]) for c in activities}
    queue = [c for c, d in indeg.items() if d == 0]
    order = []
    while queue:
        cur = queue.pop(0)
        order.append(cur)
        for nxt in succ_map[cur]:
            indeg[nxt] -= 1
            if indeg[nxt] == 0:
                queue.append(nxt)
    if len(order) != len(activities):
        return pd.DataFrame(), "Faaliyet ağında döngü tespit edildi."
    es, ef = {}, {}
    for code in order:
        es[code] = max([ef[p] for p in pred_map[code]], default=0)
        ef[code] = es[code] + activities[code]["duration"]
    proj_duration = max(ef.values(), default=0)
    lf, ls = {}, {}
    for code in reversed(order):
        lf[code] = min([ls[s] for s in succ_map[code]], default=proj_duration)
        ls[code] = lf[code] - activities[code]["duration"]
    rows = []
    for code in order:
        tf = ls[code] - es[code]
        rows.append({
            "Aktivite": code, "Adı": activities[code]["name"],
            "Süre": activities[code]["duration"],
            "ES": es[code], "EF": ef[code],
            "LS": ls[code], "LF": lf[code],
            "Bolluk": tf,
            "Kritik": "EVET" if abs(tf) < 1e-9 else "HAYIR",
        })
    return pd.DataFrame(rows), proj_duration


# ------------------------------------------------------------
# EVM
# ------------------------------------------------------------
def calculate_evm(plan, actual, progress_df):
    bac = float(plan["budget"].fillna(0).sum()) if not plan.empty else 0.0
    ac = float(actual["actual_total"].fillna(0).sum()) if not actual.empty else 0.0
    if plan.empty or progress_df.empty:
        return {
            "BAC": bac, "PV": 0.0, "EV": 0.0, "AC": ac,
            "CPI": np.nan, "SPI": np.nan, "CV": np.nan, "SV": np.nan,
            "EAC": np.nan, "ETC": np.nan, "per_wbs": pd.DataFrame(),
        }
    merged = plan[["wbs","item","category","budget"]].merge(
        progress_df, on="wbs", how="left")
    merged["actual_progress"] = merged["actual_progress"].fillna(0).clip(0,100)
    merged["planned_progress"] = merged["planned_progress"].fillna(0).clip(0,100)
    merged["PV"] = merged["budget"] * merged["planned_progress"] / 100
    merged["EV"] = merged["budget"] * merged["actual_progress"] / 100
    pv = float(merged["PV"].sum())
    ev = float(merged["EV"].sum())
    cpi = ev / ac if ac > 0 else np.nan
    spi = ev / pv if pv > 0 else np.nan
    cv = ev - ac if ac > 0 else np.nan
    sv = ev - pv if pv > 0 else np.nan
    eac = bac / cpi if (not pd.isna(cpi) and cpi > 0) else np.nan
    etc = eac - ac if not pd.isna(eac) else np.nan
    return {
        "BAC": bac, "PV": pv, "EV": ev, "AC": ac,
        "CPI": cpi, "SPI": spi, "CV": cv, "SV": sv,
        "EAC": eac, "ETC": etc,
        "per_wbs": merged[["wbs","item","category","budget",
                           "planned_progress","actual_progress",
                           "PV","EV"]].copy(),
    }


# ------------------------------------------------------------
# KALİTE SKORU
# ------------------------------------------------------------
def calculate_quality_score(plan, actual, schedule, duplicate_count,
                            open_count, estimated_count):
    score = 100
    details = []
    if duplicate_count > 0:
        d = min(20, duplicate_count * 2); score -= d
        details.append(("WBS mükerrer", duplicate_count, "Yüksek", -d))
    if open_count > 0:
        d = min(15, open_count); score -= d
        details.append(("Açık metraj", open_count, "Orta", -d))
    if estimated_count > 0:
        d = min(10, estimated_count // 2); score -= d
        details.append(("Tahmini metraj", estimated_count, "Düşük", -d))
    if actual.empty:
        score -= 25
        details.append(("Gerçek maliyet eksik", 0, "Kritik", -25))
    if schedule.empty:
        score -= 15
        details.append(("İş programı eksik", 0, "Yüksek", -15))
    if not schedule.empty and "progress" in schedule:
        missing = int(schedule["progress"].isna().sum())
        if missing > 0:
            d = min(10, missing); score -= d
            details.append(("Eksik ilerleme", missing, "Yüksek", -d))
    score = max(0, score)
    return score, pd.DataFrame(
        details, columns=["Alan","Adet","Önem","Puan Etkisi"])


# ------------------------------------------------------------
# EXCEL RAPORU
# ------------------------------------------------------------
def make_excel_report(tables):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in tables.items():
            safe = re.sub(r"[\[\]:*?/\\]", "_", sheet_name)[:31]
            if df is None:
                df = pd.DataFrame()
            if not isinstance(df, pd.DataFrame):
                df = pd.DataFrame(df)
            df.to_excel(writer, sheet_name=safe, index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            if ws.max_row > 1:
                ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="17365D")
                cell.alignment = Alignment(horizontal="center",
                                           vertical="center")
            for col in ws.columns:
                letter = col[0].column_letter
                ml = max([len(str(c.value)) for c in col
                          if c.value is not None] + [10])
                ws.column_dimensions[letter].width = min(ml + 2, 42)
    output.seek(0)
    return output.getvalue()


# ============================================================
# PREMIUM GRAFİK FONKSİYONLARI
# ============================================================

def gauge_chart(value, title, max_val=2.0, good=1.0, height=280):
    val = 0.0 if (value is None or pd.isna(value)) else float(value)
    color = BRAND["success"] if val >= good else BRAND["accent"]
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=val,
        number=dict(font=dict(size=32, color=BRAND["primary"]),
                    valueformat=".3f"),
        delta=dict(reference=good,
                   increasing=dict(color=BRAND["success"]),
                   decreasing=dict(color=BRAND["accent"])),
        title=dict(text=title, font=dict(size=13, color="#4A4A4A")),
        gauge=dict(
            axis=dict(range=[0, max_val], tickwidth=1,
                      tickcolor=BRAND["gray"], tickfont=dict(size=9)),
            bar=dict(color=color, thickness=0.72),
            bgcolor="white",
            borderwidth=0,
            steps=[
                dict(range=[0, good], color="#FFE5E7"),
                dict(range=[good, max_val], color="#E0F2EF"),
            ],
            threshold=dict(line=dict(color=BRAND["primary"], width=3),
                           thickness=0.85, value=good),
        ),
    ))
    fig.update_layout(
        height=height, margin=dict(l=20, r=20, t=50, b=10),
        paper_bgcolor="white",
        font=dict(family="Inter"),
    )
    return fig


def waterfall_budget(budget, vat, risk, height=380):
    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=["absolute", "relative", "relative", "total"],
        x=["Planlanan", "KDV", "Risk Rezervi", "Genel Toplam"],
        y=[budget, vat, risk, 0],
        text=[fmt_money(budget), fmt_money(vat),
              fmt_money(risk), fmt_money(budget + vat + risk)],
        textposition="outside",
        textfont=dict(size=12, color=BRAND["primary"]),
        connector=dict(line=dict(color=BRAND["gray"], width=1, dash="dot")),
        increasing=dict(marker=dict(color=BRAND["accent"])),
        decreasing=dict(marker=dict(color=BRAND["success"])),
        totals=dict(marker=dict(color=BRAND["primary"])),
        hovertemplate="<b>%{x}</b><br>%{text}<extra></extra>",
    ))
    return apply_theme(fig, height=height, show_legend=False,
                       title="Bütçe Yapısı — Waterfall")


def treemap_budget(plan, height=440):
    if plan.empty or "budget" not in plan:
        return None
    df = plan[["ana_grup", "category", "item", "budget"]].dropna(
        subset=["budget"]).copy()
    df = df[df["budget"] > 0]
    if df.empty:
        return None
    fig = px.treemap(
        df,
        path=[px.Constant("Tüm Proje"), "ana_grup", "category", "item"],
        values="budget",
        color="budget",
        color_continuous_scale=[
            [0.0, "#E8F4F2"], [0.3, BRAND["success"]],
            [0.6, BRAND["gold"]], [1.0, BRAND["accent"]],
        ],
    )
    fig.update_traces(
        textinfo="label+percent parent",
        textfont=dict(family="Inter", size=12),
        marker=dict(line=dict(width=2, color="white")),
        hovertemplate=("<b>%{label}</b><br>"
                       "Bütçe: ₺%{value:,.0f}<br>"
                       "Pay: %{percentParent:.1%}<extra></extra>"),
    )
    return apply_theme(fig, height=height, show_legend=False,
                       title="Bütçe Dağılımı — Treemap")


def sunburst_budget(plan, height=460):
    if plan.empty or "budget" not in plan:
        return None
    df = plan[["ana_grup", "category", "item", "budget"]].dropna(
        subset=["budget"]).copy()
    df = df[df["budget"] > 0]
    if df.empty:
        return None
    fig = px.sunburst(
        df, path=["ana_grup", "category", "item"], values="budget",
        color="budget",
        color_continuous_scale=[
            [0.0, "#E8F4F2"], [0.5, BRAND["success"]], [1.0, BRAND["accent"]],
        ],
    )
    fig.update_traces(
        textinfo="label+percent parent",
        textfont=dict(family="Inter", size=11),
        marker=dict(line=dict(width=2, color="white")),
    )
    return apply_theme(fig, height=height, show_legend=False,
                       title="Bütçe Hiyerarşisi — Sunburst")


def radar_health(evm, quality_score, actual_total, budget_total, height=420):
    cpi = 0 if pd.isna(evm["CPI"]) else min(float(evm["CPI"]), 2.0)
    spi = 0 if pd.isna(evm["SPI"]) else min(float(evm["SPI"]), 2.0)
    cost_ratio = (actual_total / budget_total) if budget_total > 0 else 0
    cost_ratio = min(cost_ratio, 2.0)
    cats = ["CPI", "SPI", "Veri Kalitesi",
            "Bütçe Kullanımı", "Risk Yönetimi"]
    vals = [cpi, spi, quality_score / 100, cost_ratio, 0.75]
    targets = [1.0, 1.0, 0.8, 1.0, 1.0]
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=targets, theta=cats, fill="toself",
        name="Hedef", line=dict(color=BRAND["success"], width=2, dash="dot"),
        fillcolor="rgba(42,157,143,0.10)",
    ))
    fig.add_trace(go.Scatterpolar(
        r=vals, theta=cats, fill="toself",
        name="Gerçek", line=dict(color=BRAND["accent"], width=3),
        fillcolor="rgba(230,57,70,0.25)",
        marker=dict(size=8, color=BRAND["accent"]),
    ))
    fig.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 2],
                            tickfont=dict(size=9),
                            gridcolor=BRAND["grid"]),
            angularaxis=dict(tickfont=dict(size=11),
                             gridcolor=BRAND["grid"]),
            bgcolor="white",
        ),
    )
    return apply_theme(fig, height=height, show_legend=True,
                       title="Proje Sağlığı — Radar")


def gantt_chart(schedule, height=520):
    if schedule.empty or "duration" not in schedule:
        return None
    df = schedule.copy()
    df["duration"] = numeric(df["duration"]).fillna(0)
    df = df[df["duration"] > 0].head(40).reset_index(drop=True)
    if df.empty:
        return None
    df["start"] = df["duration"].cumsum().shift(1).fillna(0).astype(int)
    df["end"] = df["start"] + df["duration"]
    df["label"] = df["item"].astype(str).str.slice(0, 35)

    base = pd.Timestamp("2025-01-01")
    df["date_start"] = base + pd.to_timedelta(df["start"], unit="D")
    df["date_end"] = base + pd.to_timedelta(df["end"], unit="D")

    fig = px.timeline(
        df, x_start="date_start", x_end="date_end", y="label",
        color="duration",
        color_continuous_scale=[
            [0.0, BRAND["success"]], [0.5, BRAND["warning"]],
            [1.0, BRAND["accent"]],
        ],
    )
    fig.update_yaxes(autorange="reversed")
    fig.update_traces(
        hovertemplate=("<b>%{y}</b><br>"
                       "%{x|%d.%m.%Y} → %{x2|%d.%m.%Y}<br>"
                       "Süre: %{customdata[0]} gün<extra></extra>"),
        customdata=df[["duration"]].values,
    )
    return apply_theme(fig, height=height, show_legend=False,
                       title="İş Programı — Gantt")


def sc_curve_premium(actual, budget_total, height=420):
    if actual.empty or "_date" not in actual:
        return None
    dated = actual.dropna(subset=["_date"]).sort_values("_date").copy()
    if dated.empty:
        return None
    dated["kümülatif"] = dated["actual_total"].cumsum()
    start = dated["_date"].min()
    end = dated["_date"].max()
    if pd.isna(start) or pd.isna(end) or start == end:
        return None
    total_days = (end - start).days or 1
    dated["plan_küm"] = budget_total * (dated["_date"] - start).dt.days / total_days

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=dated["_date"], y=dated["plan_küm"],
        name="Planlanan (kümülatif)",
        mode="lines",
        line=dict(color=BRAND["info"], width=2, dash="dash"),
        fill="tozeroy", fillcolor="rgba(69,123,157,0.10)",
    ))
    fig.add_trace(go.Scatter(
        x=dated["_date"], y=dated["kümülatif"],
        name="Gerçekleşen (kümülatif)",
        mode="lines+markers",
        line=dict(color=BRAND["accent"], width=3),
        marker=dict(size=7, color=BRAND["accent"],
                    line=dict(color="white", width=1.5)),
        fill="tonexty", fillcolor="rgba(230,57,70,0.15)",
    ))
    return apply_theme(fig, height=height, show_legend=True,
                       title="S-Curve — Plan vs Gerçek Kümülatif Maliyet")


def bullet_progress(categories, actuals, targets, height=400):
    fig = go.Figure()
    for i, (cat, act, tgt) in enumerate(zip(categories, actuals, targets)):
        fig.add_trace(go.Bar(
            y=[cat], x=[tgt], orientation="h",
            marker=dict(color=BRAND["grid"]),
            showlegend=False, width=0.7, hoverinfo="skip",
        ))
        fig.add_trace(go.Bar(
            y=[cat], x=[act], orientation="h",
            marker=dict(color=(BRAND["success"] if act >= tgt * 0.9
                               else BRAND["accent"])),
            name="Gerçek" if i == 0 else None,
            showlegend=(i == 0), width=0.5,
            text=[f"{act:.1f}%"], textposition="outside",
            textfont=dict(color=BRAND["primary"], size=11),
            hovertemplate=(f"<b>{cat}</b><br>"
                           f"Gerçek: {act:.1f}%<br>"
                           f"Hedef: {tgt:.1f}%<extra></extra>"),
        ))
    fig.update_layout(
        barmode="overlay", xaxis_title="Yüzde (%)",
        title="İlerleme — Hedef vs Gerçek",
    )
    return apply_theme(fig, height=height, show_legend=True)


# ============================================================
# ARAYÜZ
# ============================================================
st.title("🏗️ EMİR ASEL İNŞAAT")
st.caption(f"V{APP_VERSION} Premium | Entegre Proje, Maliyet ve Şantiye Kontrol Paneli")

with st.sidebar:
    st.header("Proje Ayarları")
    project_name = st.text_input("Proje Adı", "Villa İnşaat Projesi")
    vat_rate = st.number_input("KDV oranı", 0.0, 1.0, VAT_DEFAULT, 0.01)
    risk_rate = st.number_input("Risk rezervi", 0.0, 1.0, RISK_DEFAULT, 0.01)
    report_date = st.date_input("Raporlama tarihi", datetime.today().date())
    st.divider()
    st.caption(f"V{APP_VERSION}")

uploaded = st.file_uploader(
    "Proje kontrol Excel dosyanızı yükleyin", type=["xlsx","xlsm"])

if uploaded is None:
    st.info("Başlamak için proje Excel dosyanızı yükleyin.")
    st.markdown("""
    **V7 Premium modülleri:**
    - 🎯 Yönetici Özeti (Gauge, Treemap, Waterfall, Radar)
    - 📊 Bütçe ve Maliyet (Sunburst, Grouped Bar)
    - 📐 Metraj Kontrolü
    - 📅 İş Programı + CPM + Gantt
    - 📈 EVM + S-Curve Premium
    - 💰 Finans
    - ⚠️ Risk Matrisi
    - 📥 Excel Raporu
    """)
    st.stop()

try:
    uploaded_bytes = uploaded.getvalue()
    raw_sheets, sheet_names = load_excel(uploaded_bytes, uploaded.name)
except Exception as exc:
    st.error(f"Excel okunamadı: {exc}")
    st.stop()

plan_raw      = get_sheet(raw_sheets, "plan")
actual_raw    = get_sheet(raw_sheets, "actual")
schedule_raw  = get_sheet(raw_sheets, "schedule")
finance_raw   = get_sheet(raw_sheets, "finance")
risk_raw      = get_sheet(raw_sheets, "risk")
procurement_raw = get_sheet(raw_sheets, "procurement")

plan = prepare_plan(plan_raw)
actual = prepare_actual(actual_raw)
schedule = prepare_schedule(schedule_raw)
risk_df = prepare_risk(risk_raw)

if plan.empty:
    st.error("Planlanan maliyet / iş kalemleri sayfası bulunamadı veya boş.")
    st.write("Dosyadaki sayfalar:", sheet_names)
    st.stop()

# Ana hesaplamalar
budget_total = float(plan["budget"].fillna(0).sum())
actual_total = float(actual["actual_total"].fillna(0).sum()) if not actual.empty else 0.0
vat_budget = budget_total * vat_rate
risk_reserve = budget_total * risk_rate
budget_with_vat_risk = budget_total + vat_budget + risk_reserve

progress_df = weighted_progress(schedule)
evm = calculate_evm(plan, actual, progress_df)

plan["ana_grup"] = plan.apply(
    lambda r: category_group(r.get("category",""), r.get("item","")), axis=1)

if not actual.empty:
    plan = plan.drop(columns=["actual_total"], errors="ignore")
    actual_by_wbs = actual.groupby("wbs")["actual_total"].sum().reset_index()
    plan = plan.merge(actual_by_wbs, on="wbs", how="left")
    plan["actual_total"] = plan["actual_total"].fillna(0)
else:
    plan["actual_total"] = 0.0

if not progress_df.empty:
    plan = plan.merge(progress_df, on="wbs", how="left")

duplicate_count = int(plan["wbs"].duplicated(keep=False).sum())
status_norm = plan["status"].astype(str).str.upper()
open_count = int(status_norm.str.contains("AÇIK|ACIK", regex=True).sum())
estimated_count = int(status_norm.str.contains("TAHMİN|TAHMIN", regex=True).sum())
verified_count = int(status_norm.str.contains("DOĞRULANDI|DOGRULANDI", regex=True).sum())

quality_score, quality_details = calculate_quality_score(
    plan, actual, schedule, duplicate_count, open_count, estimated_count)

# NAVİGASYON
page = st.sidebar.radio("Modüller", [
    "🎯 Yönetici Özeti",
    "📊 Bütçe ve Maliyet",
    "📐 Metraj Kontrolü",
    "📅 İş Programı & CPM",
    "📈 EVM & S-Curve",
    "💰 Finans",
    "⚠️ Risk & Veri Kalitesi",
    "📥 Excel Raporu",
])

# ============================================================
# YÖNETİCİ ÖZETİ
# ============================================================
if page == "🎯 Yönetici Özeti":
    st.header(f"Yönetici Özeti | {project_name}")

    if quality_score >= 80:
        st.success(f"✅ Veri Kalite Skoru: {quality_score}/100 — Güvenilir")
    elif quality_score >= 60:
        st.warning(f"⚠️ Veri Kalite Skoru: {quality_score}/100 — Dikkat gerekli")
    else:
        st.error(f"❌ Veri Kalite Skoru: {quality_score}/100 — Kritik eksiklikler")

    st.divider()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Planlanan Bütçe", money(budget_total, short=True))
    c2.metric("Gerçekleşen", money(actual_total, short=True))
    c3.metric("KDV Dahil", money(budget_total + vat_budget, short=True))
    c4.metric("Risk Dahil", money(budget_with_vat_risk, short=True))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Açık Metraj", open_count)
    c6.metric("Tahmini Metraj", estimated_count)
    c7.metric("Doğrulanmış", verified_count)
    c8.metric("Mükerrer WBS", duplicate_count)

    st.divider()

    # Gauge'lar
    st.subheader("🎯 Performans Göstergeleri")
    g1, g2 = st.columns(2)
    with g1:
        st.plotly_chart(gauge_chart(evm["CPI"], "CPI — Maliyet Performansı"),
                        use_container_width=True)
    with g2:
        st.plotly_chart(gauge_chart(evm["SPI"], "SPI — Zaman Performansı"),
                        use_container_width=True)

    st.divider()

    # Treemap + Waterfall
    st.subheader("💰 Bütçe Yapısı ve Dağılımı")
    col1, col2 = st.columns([1.4, 1])
    with col1:
        tm = treemap_budget(plan)
        if tm is not None:
            st.plotly_chart(tm, use_container_width=True)
    with col2:
        st.plotly_chart(
            waterfall_budget(budget_total, vat_budget, risk_reserve),
            use_container_width=True)

    st.divider()

    # Radar
    st.subheader("🕸️ Proje Sağlığı")
    st.plotly_chart(
        radar_health(evm, quality_score, actual_total, budget_total),
        use_container_width=True)

    st.divider()

    # Grup bar + Pie
    left, right = st.columns(2)
    with left:
        st.subheader("Ana Grup Bazında Bütçe")
        grp = (plan.groupby("ana_grup")["budget"].sum()
               .reset_index().sort_values("budget", ascending=False))
        fig = px.bar(grp, x="ana_grup", y="budget",
                     color="ana_grup",
                     color_discrete_sequence=CHART_SEQUENCE,
                     text_auto=".2s")
        fig.update_traces(
            textposition="outside",
            textfont=dict(color=BRAND["primary"], size=11),
        )
        fig = apply_theme(fig, height=350, show_legend=False)
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("Planlanan vs Gerçekleşen")
        comp = pd.DataFrame({
            "Tür": ["Planlanan","Gerçekleşen"],
            "Tutar": [budget_total, actual_total],
        })
        fig = go.Figure(go.Pie(
            labels=comp["Tür"], values=comp["Tutar"], hole=0.65,
            marker=dict(colors=[BRAND["primary"], BRAND["accent"]],
                        line=dict(color="white", width=3)),
            textinfo="label+percent",
            textfont=dict(size=12, family="Inter"),
            hovertemplate="<b>%{label}</b><br>₺%{value:,.0f}<extra></extra>",
        ))
        ratio = (actual_total / budget_total * 100) if budget_total else 0
        fig.add_annotation(
            text=f"<b>{ratio:.1f}%</b><br>Gerçek/Plan",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16, color=BRAND["primary"], family="Inter"),
        )
        fig = apply_theme(fig, height=350, show_legend=True)
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.subheader("En Yüksek Bütçeli 15 Kalem")
    top = plan.nlargest(15, "budget")[
        ["wbs","item","ana_grup","budget","actual_total","status"]
    ].copy()
    top.columns = ["WBS","Kalem","Grup","Bütçe","Gerçekleşen","Statü"]
    st.dataframe(top, use_container_width=True, hide_index=True)


# ============================================================
# BÜTÇE VE MALİYET
# ============================================================
elif page == "📊 Bütçe ve Maliyet":
    st.header("Bütçe ve Maliyet Kontrolü")

    view = plan.copy()
    view["sapma"] = view["actual_total"] - view["budget"]
    view["sapma_pct"] = np.where(
        view["budget"] > 0,
        view["sapma"] / view["budget"] * 100, np.nan)

    tab1, tab2, tab3, tab4 = st.tabs(
        ["Genel Bakış", "Grup Analizi", "Hiyerarşi", "Detay Tablo"])

    with tab1:
        st.subheader("Kategori Bazında Bütçe / Gerçekleşen")
        by_cat = (view.groupby("category")[["budget","actual_total"]]
                  .sum().reset_index()
                  .sort_values("budget", ascending=False).head(15))
        fig = px.bar(
            by_cat, x="category", y=["budget","actual_total"],
            barmode="group",
            color_discrete_sequence=[BRAND["primary"], BRAND["accent"]],
            labels={"value":"Tutar (₺)","variable":"Tür"},
        )
        fig = apply_theme(fig, height=420, show_legend=True)
        fig.update_layout(xaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        st.subheader("Ana Grup Bazında Sapma")
        grp = view.groupby("ana_grup").agg(
            Bütçe=("budget","sum"),
            Gerçekleşen=("actual_total","sum"),
        ).reset_index()
        grp["Sapma"] = grp["Gerçekleşen"] - grp["Bütçe"]
        grp["Sapma %"] = np.where(grp["Bütçe"] > 0,
                                  grp["Sapma"]/grp["Bütçe"]*100, np.nan)
        st.dataframe(grp, use_container_width=True, hide_index=True)

        fig = px.bar(
            grp, x="ana_grup", y="Sapma",
            color="Sapma",
            color_continuous_scale=[
                [0.0, BRAND["success"]],
                [0.5, BRAND["gold"]],
                [1.0, BRAND["accent"]],
            ],
            text="Sapma",
        )
        fig.update_traces(
            texttemplate="₺%{text:,.0f}",
            textposition="outside",
            textfont=dict(color=BRAND["primary"], size=11),
        )
        fig = apply_theme(fig, height=380, show_legend=False)
        st.plotly_chart(fig, use_container_width=True)

    with tab3:
        st.subheader("Bütçe Hiyerarşisi — Sunburst")
        sb = sunburst_budget(plan)
        if sb is not None:
            st.plotly_chart(sb, use_container_width=True)
        else:
            st.info("Sunburst için pozitif bütçeli kalem gerekli.")

    with tab4:
        st.subheader("Tüm Kalemler")
        disp = view[["wbs","item","category","budget","actual_total",
                     "sapma","sapma_pct","status"]].copy()
        disp.columns = ["WBS","Kalem","Kategori","Bütçe","Gerçekleşen",
                        "Sapma","Sapma %","Statü"]
        st.dataframe(disp, use_container_width=True, hide_index=True)

    st.caption("Gerçek maliyet, fiili kayıtların toplamıdır. "
               "Sipariş tek başına gerçek maliyet sayılmaz.")


# ============================================================
# METRAJ KONTROLÜ
# ============================================================
elif page == "📐 Metraj Kontrolü":
    st.header("Metraj ve Miktar Kontrolü")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Toplam Kalem", len(plan))
    c2.metric("Açık", open_count)
    c3.metric("Tahmini", estimated_count)
    c4.metric("Doğrulanmış", verified_count)

    st.divider()

    left, right = st.columns(2)
    with left:
        st.subheader("Statü Dağılımı")
        sc = plan["status"].value_counts().reset_index()
        sc.columns = ["Durum","Adet"]
        fig = px.pie(sc, names="Durum", values="Adet", hole=0.55,
                     color_discrete_sequence=CHART_SEQUENCE)
        fig.update_traces(
            textinfo="label+percent",
            textfont=dict(size=11),
            marker=dict(line=dict(color="white", width=2)),
        )
        fig = apply_theme(fig, height=380, show_legend=True)
        st.plotly_chart(fig, use_container_width=True)

    with right:
        st.subheader("Statü Bazında Bütçe")
        bs = plan.groupby("status")["budget"].sum().reset_index()
        fig = px.bar(bs, x="status", y="budget",
                     color="status",
                     color_discrete_sequence=CHART_SEQUENCE,
                     text_auto=".2s")
        fig.update_traces(
            textposition="outside",
            textfont=dict(color=BRAND["primary"], size=11),
        )
        fig = apply_theme(fig, height=380, show_legend=False)
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.subheader("Açık ve Tahmini Metrajlar")
    pending = plan[plan["status"].str.upper().str.contains(
        "AÇIK|ACIK|TAHMİN|TAHMIN", regex=True)]
    if pending.empty:
        st.success("✅ Açık/tahmini metraj yok.")
    else:
        disp = pending[["wbs","item","unit","qty","status","category"]].copy()
        disp.columns = ["WBS","Kalem","Birim","Miktar","Statü","Kategori"]
        st.dataframe(disp, use_container_width=True, hide_index=True)
        st.warning(
            f"⚠️ {len(pending)} kalem onaylı proje metrajı değil. "
            "Statik donatı listesi, mimari çizimler ve MEP projeleriyle doğrulanmalı.")


# ============================================================
# İŞ PROGRAMI & CPM
# ============================================================
elif page == "📅 İş Programı & CPM":
    st.header("İş Programı & Kritik Yol")

    if schedule.empty:
        st.warning("İş programı sayfası bulunamadı.")
        st.stop()

    tab1, tab2, tab3, tab4 = st.tabs(
        ["Aktiviteler", "CPM Analizi", "Gantt", "İlerleme Karşılaştırma"])

    with tab1:
        st.subheader("Aktivite Listesi")
        st.dataframe(schedule, use_container_width=True, hide_index=True)

    with tab2:
        st.subheader("Kritik Yol Analizi")
        cpm_df, result = calculate_cpm(schedule)
        if isinstance(result, str):
            st.error(f"❌ {result}")
        elif cpm_df.empty:
            st.info("CPM için yeterli veri yok.")
        else:
            st.success(f"✅ Proje CPM Süresi: **{result:.0f} gün**")
            kc = int((cpm_df["Kritik"] == "EVET").sum())
            c1, c2 = st.columns(2)
            c1.metric("Kritik Aktivite", kc)
            c2.metric("Toplam Aktivite", len(cpm_df))
            st.dataframe(cpm_df, use_container_width=True, hide_index=True)
            st.subheader("Sadece Kritik Aktiviteler")
            st.dataframe(cpm_df[cpm_df["Kritik"] == "EVET"],
                         use_container_width=True, hide_index=True)

    with tab3:
        st.subheader("Gantt Şeması")
        gt = gantt_chart(schedule)
        if gt is None:
            st.info("Gantt için süre verisi gerekli.")
        else:
            st.plotly_chart(gt, use_container_width=True)

    with tab4:
        st.subheader("Planlanan vs Gerçekleşen İlerleme")
        valid = schedule.dropna(subset=["progress"]).copy()
        if valid.empty:
            st.info("İlerleme verisi yok.")
        else:
            planned_vals = (valid["planned_progress"]
                            if "planned_progress" in valid
                            else pd.Series([0]*len(valid)))
            fig = go.Figure()
            fig.add_trace(go.Bar(
                name="Planlanan %",
                x=valid["wbs"].head(30),
                y=planned_vals.head(30),
                marker_color=BRAND["primary"],
            ))
            fig.add_trace(go.Bar(
                name="Gerçekleşen %",
                x=valid["wbs"].head(30),
                y=valid["progress"].head(30),
                marker_color=BRAND["accent"],
            ))
            fig.update_layout(barmode="group", xaxis_title="",
                              yaxis_title="İlerleme (%)")
            fig = apply_theme(fig, height=420, show_legend=True,
                              title="Aktivite Bazında Plan vs Gerçek")
            st.plotly_chart(fig, use_container_width=True)


# ============================================================
# EVM & S-CURVE
# ============================================================
elif page == "📈 EVM & S-Curve":
    st.header("Kazanılmış Değer Yönetimi & S-Curve")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("PV (Planlanan)", money(evm['PV'], short=True))
    c2.metric("EV (Kazanılmış)", money(evm['EV'], short=True))
    c3.metric("AC (Gerçekleşen)", money(evm['AC'], short=True))
    c4.metric("BAC (Bütçe)", money(evm['BAC'], short=True))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("CPI", f"{evm['CPI']:.3f}" if not pd.isna(evm['CPI']) else "—")
    c6.metric("SPI", f"{evm['SPI']:.3f}" if not pd.isna(evm['SPI']) else "—")
    c7.metric("CV", money(evm['CV'], short=True) if not pd.isna(evm['CV']) else "—")
    c8.metric("SV", money(evm['SV'], short=True) if not pd.isna(evm['SV']) else "—")

    st.divider()

    tab1, tab2, tab3 = st.tabs(["EVM Grafiği", "S-Curve", "Detay"])

    with tab1:
        cats = ["BAC","PV","EV","AC"]
        vals = [evm["BAC"], evm["PV"], evm["EV"], evm["AC"]]
        fig = go.Figure(data=[go.Bar(
            x=cats, y=vals,
            text=[fmt_money(v) for v in vals],
            textposition="outside",
            textfont=dict(color=BRAND["primary"], size=12),
            marker=dict(color=[BRAND["primary"], BRAND["info"],
                               BRAND["success"], BRAND["accent"]]),
        )])
        fig.update_layout(yaxis_title="Tutar")
        fig = apply_theme(fig, height=420, show_legend=False,
                          title="EVM Değerleri")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Yorumlama")
        if not pd.isna(evm['CPI']) and evm['CPI'] < 1:
            st.warning(f"⚠️ CPI = {evm['CPI']:.3f} < 1 — Maliyet performansı olumsuz.")
        elif not pd.isna(evm['CPI']):
            st.success(f"✅ CPI = {evm['CPI']:.3f} ≥ 1 — Maliyet performansı olumlu.")
        if not pd.isna(evm['SPI']) and evm['SPI'] < 1:
            st.warning(f"⚠️ SPI = {evm['SPI']:.3f} < 1 — Proje planın gerisinde.")
        elif not pd.isna(evm['SPI']):
            st.success(f"✅ SPI = {evm['SPI']:.3f} ≥ 1 — Proje planla uyumlu.")

        # Bullet chart
        st.subheader("Kategori Bazında İlerleme — Hedef vs Gerçek")
        if not evm["per_wbs"].empty:
            pw = (evm["per_wbs"]
                  .groupby("category")[["planned_progress","actual_progress"]]
                  .mean().reset_index()
                  .sort_values("planned_progress", ascending=False)
                  .head(10))
            fig_b = bullet_progress(
                pw["category"].tolist(),
                pw["actual_progress"].fillna(0).tolist(),
                pw["planned_progress"].fillna(0).tolist(),
            )
            st.plotly_chart(fig_b, use_container_width=True)

    with tab2:
        st.subheader("S-Curve Premium")
        sc = sc_curve_premium(actual, budget_total)
        if sc is None:
            st.info("S-Curve için tarihli gerçek maliyet kaydı gerekli.")
        else:
            st.plotly_chart(sc, use_container_width=True)

    with tab3:
        st.subheader("WBS Bazında EVM Detayı")
        if evm["per_wbs"].empty:
            st.info("Veri yok.")
        else:
            detay = evm["per_wbs"].copy()
            detay.columns = ["WBS","Kalem","Kategori","Bütçe",
                             "Plan %","Gerçek %","PV","EV"]
            st.dataframe(detay, use_container_width=True, hide_index=True)


# ============================================================
# FİNANS
# ============================================================
elif page == "💰 Finans":
    st.header("Finans ve Nakit Akışı")

    if finance_raw.empty:
        st.warning("Finans sayfası bulunamadı.")
        st.stop()

    st.dataframe(finance_raw, use_container_width=True, hide_index=True)

    date_col = find_col(finance_raw, "date")
    amount_col = find_col(finance_raw, "amount")
    direction_col = find_col(finance_raw, "direction")

    if date_col and amount_col:
        f = finance_raw.copy()
        f["_date"] = pd.to_datetime(f[date_col], errors="coerce", dayfirst=True)
        f["_amount"] = numeric(f[amount_col])
        f = f.dropna(subset=["_date","_amount"])

        if not f.empty:
            f["Ay"] = f["_date"].dt.to_period("M").astype(str)
            if direction_col:
                f["_direction"] = f[direction_col].astype(str).str.lower()
                f["_signed"] = np.where(
                    f["_direction"].str.contains("ödeme|odeme|çıkış|cikis|out"),
                    -f["_amount"].abs(), f["_amount"].abs())
            else:
                f["_signed"] = f["_amount"]

            monthly = f.groupby("Ay").agg(
                Net=("_signed","sum"),
                Giriş=("_amount", lambda x: x[x>0].sum()),
                Çıkış=("_amount", lambda x: abs(x[x<0]).sum()),
            ).reset_index()

            fig = go.Figure()
            fig.add_trace(go.Bar(x=monthly["Ay"], y=monthly["Giriş"],
                                 name="Giriş", marker_color=BRAND["success"]))
            fig.add_trace(go.Bar(x=monthly["Ay"], y=monthly["Çıkış"],
                                 name="Çıkış", marker_color=BRAND["accent"]))
            fig.add_trace(go.Scatter(x=monthly["Ay"], y=monthly["Net"],
                                     mode="lines+markers", name="Net",
                                     line=dict(color=BRAND["primary"], width=3)))
            fig.update_layout(barmode="group",
                              xaxis_title="Ay", yaxis_title="Tutar (₺)")
            fig = apply_theme(fig, height=420, show_legend=True,
                              title="Aylık Nakit Hareketi")
            st.plotly_chart(fig, use_container_width=True)


# ============================================================
# RİSK & VERİ KALİTESİ
# ============================================================
elif page == "⚠️ Risk & Veri Kalitesi":
    st.header("Risk Yönetimi ve Veri Kalitesi")
    tab1, tab2 = st.tabs(["🎯 Risk Matrisi", "📋 Veri Kalitesi"])

    with tab1:
        if risk_df.empty:
            st.info("Risk kayıt sayfası bulunamadı.")
        else:
            c1, c2, c3 = st.columns(3)
            c1.metric("Toplam Risk", len(risk_df))
            c2.metric("Yüksek/Çok Yüksek",
                      int(risk_df["_risk_level"]
                          .isin(["YÜKSEK","ÇOK YÜKSEK"]).sum()))
            c3.metric("En Yüksek Puan",
                      f"{risk_df['_risk_score'].max():.1f}")

            st.divider()
            st.subheader("Risk Isı Haritası")

            pivot = risk_df.pivot_table(
                index="_impact", columns="_probability",
                values="_risk_score", aggfunc="count", fill_value=0)

            if not pivot.empty:
                fig = px.imshow(
                    pivot, text_auto=True,
                    color_continuous_scale=[
                        [0.0, "#E0F2EF"],
                        [0.4, BRAND["gold"]],
                        [1.0, BRAND["accent"]],
                    ],
                    labels=dict(x="Olasılık", y="Etki", color="Risk Sayısı"),
                )
                fig.update_traces(
                    textfont=dict(size=14, color=BRAND["primary"]),
                )
                fig = apply_theme(fig, height=420, show_legend=False,
                                  title="Olasılık × Etki — Isı Haritası")
                st.plotly_chart(fig, use_container_width=True)

            st.subheader("Risk Listesi")
            disp = risk_df.copy()
            if "Kalem" not in disp and "Konu" in disp:
                disp["Kalem"] = disp["Konu"]
            st.dataframe(disp, use_container_width=True, hide_index=True)

    with tab2:
        st.subheader(f"Veri Kalite Skoru: **{quality_score}/100**")
        if quality_score >= 80:
            st.success("✅ Veri kalitesi yüksek")
        elif quality_score >= 60:
            st.warning("⚠️ Orta düzey — bazı iyileştirmeler gerekli")
        else:
            st.error("❌ Düşük — kritik eksiklikler var")
        st.progress(float(quality_score) / 100.0)

        st.divider()
        st.subheader("Detaylı Bulgular")
        if quality_details.empty:
            st.success("✅ Otomatik kontrollerde bulgu yok.")
        else:
            st.dataframe(quality_details, use_container_width=True,
                         hide_index=True)

        st.subheader("Kontrol Kuralları")
        st.markdown("""
        - ✅ Donatı miktarları statik proje ve BBS ile doğrulanmalı
        - ✅ Teklif fiyatları tarih + tedarikçi + KDV durumuyla saklanmalı
        - ✅ Fatura/hakediş mükerrerlik kontrolü yapılmalı
        - ✅ Sipariş ile fiili maliyet ayrı izlenmeli
        - ✅ Risk rezervi gerçekleşen maliyet değildir
        """)


# ============================================================
# EXCEL RAPORU
# ============================================================
elif page == "📥 Excel Raporu":
    st.header("Excel Raporu")
    st.markdown("""
    **Bu modül orijinal Excel dosyanızı değiştirmez.** Yeni bir rapor dosyası oluşturur.
    """)

    report_tables = {
        "Yonetici_Ozeti": pd.DataFrame({
            "Gösterge": ["Proje","Rapor Tarihi","Planlanan Bütçe","KDV",
                         "Risk Rezervi","KDV+Risk Dahil","Gerçekleşen",
                         "PV","EV","AC","CPI","SPI","Veri Kalite Skoru"],
            "Değer": [project_name, str(report_date), budget_total,
                      vat_budget, risk_reserve, budget_with_vat_risk,
                      actual_total, evm["PV"], evm["EV"], evm["AC"],
                      evm["CPI"] if not pd.isna(evm["CPI"]) else "—",
                      evm["SPI"] if not pd.isna(evm["SPI"]) else "—",
                      quality_score],
        }),
        "Butce_Gerceklesen": plan,
        "Metraj_Kontrol": plan[["wbs","item","category","unit","qty","status"]],
        "Gercek_Maliyet": actual,
        "Is_Programi": schedule,
        "Finans": finance_raw,
        "Risk": risk_df,
        "Veri_Kalite": quality_details,
        "EVM_Detay": evm["per_wbs"],
    }

    if st.button("📥 Rapor Oluştur", type="primary"):
        try:
            excel_bytes = make_excel_report(report_tables)
            st.download_button(
                "💾 Excel İndir",
                data=excel_bytes,
                file_name=f"Emir_Asel_Rapor_{report_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.success("✅ Rapor hazır. İndirme butonunu kullanın.")
        except Exception as exc:
            st.error(f"Rapor hatası: {exc}")


# ============================================================
# ALT BİLGİ
# ============================================================
st.divider()
st.caption(f"Emir Asel İnşaat | Proje Kontrol V{APP_VERSION} Premium | "
           f"Son: {datetime.now():%d.%m.%Y %H:%M}")
