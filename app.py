# -*- coding: utf-8 -*-
"""
EMİR ASEL İNŞAAT | PROJE KONTROL V8.2 (FİNAL)
- V8.1 hesap motoru (CPM, EVM, weighted_progress) birebir korundu
- Kurumsal tema + audit + snapshot + metraj/hakediş + stok + S-curve + PDF/JSON/Excel
Çalıştırma: streamlit run emir_asel_v8_2.py
"""
from __future__ import annotations
from io import BytesIO
from datetime import date, datetime
import re, json
import sqlite3
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors as rl_colors
    from reportlab.lib.units import cm
    PDF_OK = True
except ImportError:
    PDF_OK = False

try:
    from streamlit_option_menu import option_menu
    OPTION_MENU_OK = True
except ImportError:
    OPTION_MENU_OK = False

APP_VERSION = "8.2"
VAT_DEFAULT = 0.20

st.set_page_config(
    page_title="Emir Asel | Proje Kontrol V8.2",
    page_icon="🏗️", layout="wide", initial_sidebar_state="expanded",
)

# ============================================================
# TEMA
# ============================================================
NAVY = "#14263D"; NAVY_2 = "#0A192F"
RED = "#C83E4D"; RED_ACCENT = "#E63946"
TEAL = "#278C82"; GOLD = "#D6A84B"; LIGHT = "#F2F5F8"
PALETTE = [NAVY, RED, TEAL, GOLD, "#6487A5", "#8D79A8", "#7D9B75"]

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap');
    html, body, [class*="css"] {{ font-family: 'Montserrat', sans-serif; }}
    .stApp {{ background: linear-gradient(180deg, #F8F9FA 0%, #FFFFFF 100%); }}
    .ea-header {{
        background: linear-gradient(135deg, {NAVY_2} 0%, {NAVY} 100%);
        padding: 26px 32px; border-radius: 16px; color: white;
        margin-bottom: 22px; box-shadow: 0 8px 24px rgba(10,25,47,0.22);
        position: relative; overflow: hidden;
    }}
    .ea-header::before {{
        content: ""; position: absolute; top: -50%; right: -10%;
        width: 360px; height: 360px;
        background: radial-gradient(circle, rgba(200,62,77,0.18) 0%, transparent 70%);
        border-radius: 50%;
    }}
    .ea-header h1 {{
        color: white !important; margin: 0;
        font-size: 26px; font-weight: 800; letter-spacing: 1.2px;
        text-transform: uppercase; position: relative;
    }}
    .ea-header .sub {{
        color: #B8C5D6; font-size: 12px; margin-top: 6px;
        letter-spacing: 2px; font-weight: 500;
    }}
    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {NAVY_2} 0%, {NAVY} 100%);
    }}
    section[data-testid="stSidebar"] * {{ color: white !important; }}
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {{ font-weight: 700; }}
    div[data-testid="stMetric"] {{
        background: white; border-radius: 12px; padding: 16px 20px;
        box-shadow: 0 3px 10px rgba(10,25,47,0.06);
        border-left: 5px solid {RED};
    }}
    div[data-testid="stMetric"] label {{
        color: #64748B !important; font-weight: 600 !important;
        font-size: 11px !important; text-transform: uppercase; letter-spacing: 0.5px;
    }}
    div[data-testid="stMetricValue"] {{
        color: {NAVY} !important; font-weight: 800 !important; font-size: 22px !important;
    }}
    .stButton > button {{
        background: linear-gradient(135deg, {RED_ACCENT} 0%, #A42A38 100%);
        color: white; border: none; border-radius: 10px;
        padding: 10px 20px; font-weight: 700;
        box-shadow: 0 4px 12px rgba(200,62,77,0.22);
    }}
    .section-title {{
        color: {NAVY}; font-size: 18px; font-weight: 800;
        margin: 22px 0 10px 0; padding-left: 12px;
        border-left: 4px solid {RED_ACCENT};
    }}
    .badge-saved {{
        display: inline-block; padding: 5px 12px; border-radius: 12px;
        font-size: 11px; font-weight: 700; background: #E8F5E9;
        color: #2E7D32; border: 1px solid #A5D6A7;
    }}
    .badge-db {{
        display: inline-block; padding: 5px 12px; border-radius: 12px;
        font-size: 11px; font-weight: 700; background: #EFF6FF;
        color: #1976D2; border: 1px solid #93C5FD;
    }}
    .badge-warn {{
        display: inline-block; padding: 5px 12px; border-radius: 12px;
        font-size: 11px; font-weight: 700; background: #FEF3C7;
        color: #92400E; border: 1px solid #FCD34D;
    }}
    .ea-footer {{
        margin-top: 40px; padding: 20px 0; text-align: center;
        color: #94A3B8; font-size: 11px; border-top: 1px solid #E5EAF0;
    }}
</style>
""", unsafe_allow_html=True)

# ============================================================
# ŞEMALAR
# ============================================================
PLAN_COLS = ["WBS","Poz No","İş Kalemi","Ana Kategori","Alt Kategori",
             "Birim","Miktar","Birim Fiyat","Planlanan Tutar",
             "Başlangıç","Bitiş","Sorumlu","Açıklama"]
ACTUAL_COLS = ["WBS","Poz No","İş Kalemi","Ana Kategori","Birim",
               "Gerçekleşen Miktar","Gerçekleşen Birim Fiyat",
               "Gerçekleşen Tutar","Tarih","Tedarikçi","Belge No","Açıklama"]
SCHEDULE_COLS = ["WBS","Aktivite","Süre (gün)","Öncül WBS","Bağımlılık Tipi",
                 "Başlangıç","Bitiş","Fiili Başlangıç","Fiili Bitiş",
                 "Tamamlanma (%)","Ağırlık (%)","Sorumlu","Durum","Kanıt / Saha Notu"]
RISK_COLS = ["Risk ID","Risk Tanımı","Kategori","Olasılık (1-5)",
             "Etki (1-5)","Risk Skoru","Önlem","Sorumlu","Durum",
             "Tahmini Maliyet Etkisi"]

COST_CATEGORIES = ["Mimarlık / Proje Gideri","Yapı Denetim Gideri",
    "Yapı Analiz Laboratuvar Gideri","Hafriyat","Grobeton ve Beton",
    "Toplam Kalıp + Demir İşçiliği","Demir / Donatı Malzemesi",
    "Elektrik + Su İşçiliği","Elektrik Malzemeleri","Mekanik Tesisat",
    "Duvar ve Sıva","Yalıtım","Şap ve Zemin","Seramik ve Kaplama",
    "Kapı / Pencere / Doğrama","Cephe","Çatı","Boya",
    "Peyzaj ve Çevre Düzenleme","Nakliye / Lojistik",
    "Şantiye Genel Giderleri","İş Makinesi / Ekipman","İSG / Çevre","Diğer"]

# ============================================================
# YARDIMCI
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

def normalize_dates(df, columns):
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
    return df

def ensure_columns(df, columns):
    out = df.copy()
    for c in columns:
        if c not in out.columns: out[c] = np.nan
    return out[columns]

def money(v, currency="₺"):
    try:
        return f"{currency} {float(v):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    except Exception:
        return f"{currency} 0,00"

def pct(v):
    return f"%{float(v):.1f}".replace(".",",")

def calc_plan_amount(df):
    df = df.copy()
    q = pd.to_numeric(df["Miktar"], errors="coerce").fillna(0)
    p = pd.to_numeric(df["Birim Fiyat"], errors="coerce").fillna(0)
    df["Planlanan Tutar"] = q * p
    return df

def calc_actual_amount(df):
    df = df.copy()
    q = pd.to_numeric(df["Gerçekleşen Miktar"], errors="coerce").fillna(0)
    p = pd.to_numeric(df["Gerçekleşen Birim Fiyat"], errors="coerce").fillna(0)
    df["Gerçekleşen Tutar"] = q * p
    return df

def summarize_costs(plan, actual):
    p = plan.groupby("Ana Kategori", dropna=False)["Planlanan Tutar"].sum().rename("Bütçe")
    a = actual.groupby("Ana Kategori", dropna=False)["Gerçekleşen Tutar"].sum().rename("Gerçekleşen")
    result = pd.concat([p, a], axis=1).fillna(0).reset_index()
    result["Sapma"] = result["Gerçekleşen"] - result["Bütçe"]
    result["Sapma (%)"] = np.where(result["Bütçe"]!=0, result["Sapma"]/result["Bütçe"]*100, np.nan)
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
# CPM MOTORU
# ============================================================
def calculate_cpm(schedule):
    if schedule.empty: return pd.DataFrame(), []
    df = schedule.copy().reset_index(drop=True)
    df["WBS"] = df["WBS"].astype(str).str.strip()
    df["Süre (gün)"] = pd.to_numeric(df["Süre (gün)"], errors="coerce").fillna(0).clip(lower=0)
    duration = dict(zip(df["WBS"], df["Süre (gün)"].astype(float)))
    preds, successors = {}, {w: [] for w in duration}
    indegree = {w: 0 for w in duration}
    warnings = []
    if df["WBS"].duplicated().any():
        warnings.append("Tekrarlanan WBS kodu var; CPM güvenilir olmayabilir.")
    for _, row in df.iterrows():
        task = str(row["WBS"]).strip()
        default_rel = str(row.get("Bağımlılık Tipi","FS") or "FS").upper().strip()
        if default_rel not in {"FS","SS","FF","SF"}:
            warnings.append(f"{task}: geçersiz varsayılan ilişki '{default_rel}', FS kabul edildi.")
            default_rel = "FS"
        raw = "" if pd.isna(row.get("Öncül WBS")) else str(row.get("Öncül WBS"))
        edges = []
        for token in re.split(r"[,;]", raw):
            token = token.strip()
            if not token or token.lower()=="nan": continue
            parts = re.split(r"[:=]", token, maxsplit=1)
            parent = parts[0].strip()
            rel = parts[1].strip().upper() if len(parts)==2 else default_rel
            if rel not in {"FS","SS","FF","SF"}:
                warnings.append(f"{task}: '{token}' tipi geçersiz; {default_rel} kullanıldı.")
                rel = default_rel
            edges.append((parent, rel))
        preds[task] = edges
    for task, edges in preds.items():
        for parent, rel in edges:
            if parent not in duration:
                warnings.append(f"{task}: öncül WBS '{parent}' bulunamadı.")
                continue
            successors[parent].append((task, rel))
            indegree[task] += 1
    queue = [w for w,d in indegree.items() if d==0]
    order = []
    while queue:
        node = queue.pop(0); order.append(node)
        for nxt,_ in successors[node]:
            indegree[nxt] -= 1
            if indegree[nxt]==0: queue.append(nxt)
    if len(order)!=len(duration):
        warnings.append("Döngüsel öncül ilişkisi. CPM hesaplanamadı.")
        return df.assign(**{"Erken Başlangıç":np.nan,"Erken Bitiş":np.nan,
            "Geç Başlangıç":np.nan,"Geç Bitiş":np.nan,"Bolluk":np.nan,
            "Kritik mi?":"Kontrol gerekli"}), warnings
    es, ef = {}, {}
    for task in order:
        d = duration[task]; constraints = []
        for parent, rel in preds[task]:
            if parent not in duration: continue
            if rel=="FS": constraints.append(ef[parent]+1)
            elif rel=="SS": constraints.append(es[parent])
            elif rel=="FF": constraints.append(ef[parent]-d+1)
            elif rel=="SF": constraints.append(es[parent]-d+1)
        es[task] = max([1]+constraints); ef[task] = es[task]+d-1
    project_duration = max(ef.values(), default=0)
    ls, lf = {}, {}
    for task in reversed(order):
        d = duration[task]; latest_finish_constraints = []
        for nxt, rel in successors[task]:
            if rel=="FS": latest_finish_constraints.append(ls[nxt]-1)
            elif rel=="SS": latest_finish_constraints.append(ls[nxt]+d-1)
            elif rel=="FF": latest_finish_constraints.append(lf[nxt])
            elif rel=="SF": latest_finish_constraints.append(lf[nxt]+d-1)
        lf[task] = min(latest_finish_constraints) if latest_finish_constraints else project_duration
        ls[task] = lf[task]-d+1
    df["Erken Başlangıç"] = df["WBS"].map(es)
    df["Erken Bitiş"] = df["WBS"].map(ef)
    df["Geç Başlangıç"] = df["WBS"].map(ls)
    df["Geç Bitiş"] = df["WBS"].map(lf)
    df["Bolluk"] = df["WBS"].map(ls) - df["WBS"].map(es)
    df["Kritik mi?"] = np.where(df["Bolluk"].abs()<1e-8, "EVET", "Hayır")
    return df, warnings

# ============================================================
# EVM MOTORU
# ============================================================
def calculate_evm(plan, actual, schedule):
    bac = float(plan["Planlanan Tutar"].sum())
    ac = float(actual["Gerçekleşen Tutar"].sum())
    pv = 0.0; ev = 0.0
    report_date = pd.Timestamp(date.today())
    if not plan.empty and not schedule.empty:
        sch = schedule.copy()
        sch["WBS"] = sch["WBS"].astype(str).str.strip()
        sch["Süre (gün)"] = pd.to_numeric(sch["Süre (gün)"], errors="coerce").fillna(0).clip(lower=0)
        sch["Tamamlanma (%)"] = pd.to_numeric(sch["Tamamlanma (%)"], errors="coerce").fillna(0).clip(0,100)
        sch["Başlangıç"] = pd.to_datetime(sch["Başlangıç"], errors="coerce", dayfirst=True)
        sch["Bitiş"] = pd.to_datetime(sch["Bitiş"], errors="coerce", dayfirst=True)
        progress_by_wbs, planned_by_wbs = {}, {}
        for wbs, group in sch.groupby("WBS"):
            if "Ağırlık (%)" in group.columns:
                weights = pd.to_numeric(group["Ağırlık (%)"], errors="coerce").fillna(0).clip(lower=0)
            else:
                weights = pd.Series(0.0, index=group.index)
            if weights.sum()<=0: weights = group["Süre (gün)"].replace(0,np.nan).fillna(0)
            if weights.sum()>0:
                physical = float(np.average(group["Tamamlanma (%)"], weights=weights))
            else:
                physical = float(group["Tamamlanma (%)"].mean())
            progress_by_wbs[wbs] = min(max(physical/100.0,0.0),1.0)
            fractions = []
            for _, task in group.iterrows():
                st_date, end_date = task["Başlangıç"], task["Bitiş"]
                if pd.isna(st_date) or pd.isna(end_date) or end_date<=st_date: continue
                if report_date<=st_date: fraction = 0.0
                elif report_date>=end_date: fraction = 1.0
                else: fraction = (report_date-st_date).total_seconds()/(end_date-st_date).total_seconds()
                fractions.append((fraction, max(float(task["Süre (gün)"]),1.0)))
            if fractions:
                planned_by_wbs[wbs] = float(np.average([x[0] for x in fractions], weights=[x[1] for x in fractions]))
        for _, row in plan.iterrows():
            wbs = str(row["WBS"]).strip()
            amount = safe_num(row["Planlanan Tutar"])
            ev += amount * progress_by_wbs.get(wbs,0.0)
            pv += amount * planned_by_wbs.get(wbs,0.0)
    cpi = ev/ac if ac>0 else np.nan
    spi = ev/pv if pv>0 else np.nan
    cv = ev-ac; sv = ev-pv
    eac = bac/cpi if pd.notna(cpi) and cpi>0 else np.nan
    return {"BAC":bac,"PV":pv,"EV":ev,"AC":ac,"CV":cv,"SV":sv,
            "CPI":cpi,"SPI":spi,"EAC":eac,
            "VAC": bac-eac if pd.notna(eac) else np.nan}

def excel_export(tables):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, df in tables.items():
            safe_sheet = str(sheet)[:31]
            df.to_excel(writer, sheet_name=safe_sheet, index=False)
            ws = writer.sheets[safe_sheet]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="14263D")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col_idx, col in enumerate(df.columns, 1):
                values = [str(col)] + [str(x)[:60] for x in df[col].head(200).fillna("")]
                width = min(max(max(map(len, values), default=10)+2, 12), 38)
                ws.column_dimensions[get_column_letter(col_idx)].width = width
            ws.row_dimensions[1].height = 32
    output.seek(0)
    return output.getvalue()

# ============================================================
# SQLite
# ============================================================
DB_PATH = Path(__file__).with_name("emir_asel_v8.db")

def db_connect():
    con = sqlite3.connect(DB_PATH, timeout=20)
    con.row_factory = sqlite3.Row
    return con

def init_local_db():
    with db_connect() as con:
        con.execute("""CREATE TABLE IF NOT EXISTS purchase_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT, project TEXT, order_no TEXT, item TEXT,
            supplier TEXT, quantity REAL, unit TEXT, unit_price REAL, total REAL,
            order_date TEXT, due_date TEXT, status TEXT, delivered_qty REAL DEFAULT 0,
            invoiced_amount REAL DEFAULT 0, paid_amount REAL DEFAULT 0, wbs TEXT, note TEXT)""")
        con.execute("""CREATE TABLE IF NOT EXISTS labor_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT, project TEXT, work_date TEXT, team TEXT,
            work_description TEXT, worker_count REAL, daily_wage REAL, day_count REAL,
            total REAL, wbs TEXT, subcontractor TEXT, note TEXT)""")
        con.execute("""CREATE TABLE IF NOT EXISTS progress_records (
            project TEXT, wbs TEXT, progress REAL DEFAULT 0, evidence TEXT,
            updated_at TEXT, PRIMARY KEY(project, wbs))""")
        con.execute("""CREATE TABLE IF NOT EXISTS stock_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT, project TEXT, material_name TEXT,
            category TEXT, unit TEXT, quantity_in REAL DEFAULT 0, quantity_out REAL DEFAULT 0,
            unit_price REAL DEFAULT 0, supplier TEXT, update_date TEXT)""")
        con.execute("""CREATE TABLE IF NOT EXISTS cashflow_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT, project TEXT, ay TEXT,
            nakit_akisi REAL DEFAULT 0, UNIQUE(project, ay))""")
        con.execute("""CREATE TABLE IF NOT EXISTS project_snapshots (
            project TEXT PRIMARY KEY, payload BLOB, updated_at TEXT)""")
        con.execute("""CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT, project TEXT, tablo TEXT,
            kayit_id TEXT, alan TEXT, eski TEXT, yeni TEXT,
            kullanici TEXT DEFAULT 'local_user',
            tarih TEXT DEFAULT (datetime('now','localtime')))""")

def db_read(table, project):
    allowed = {"purchase_records","labor_records"}
    if table not in allowed: return pd.DataFrame()
    with db_connect() as con:
        return pd.read_sql_query(f"SELECT * FROM {table} WHERE project = ? ORDER BY id DESC",
                                 con, params=(project,))

def db_save_frame(table, frame, project):
    if table not in {"purchase_records","labor_records"}: raise ValueError("Geçersiz tablo")
    with db_connect() as con:
        con.execute(f"DELETE FROM {table} WHERE project = ?", (project,))
        if not frame.empty:
            payload = frame.copy(); payload["project"] = project
            cols = [c for c in payload.columns if c != "id"]
            payload[cols].to_sql(table, con, if_exists="append", index=False)

def db_load_progress(project):
    with db_connect() as con:
        rows = con.execute("SELECT wbs, progress, evidence FROM progress_records WHERE project = ?",
                           (project,)).fetchall()
    return {r["wbs"]:{"progress":r["progress"],"evidence":r["evidence"]} for r in rows}

def db_save_progress(project, schedule_frame):
    with db_connect() as con:
        for _, row in schedule_frame.iterrows():
            wbs = str(row.get("WBS","")).strip()
            if not wbs or wbs.lower()=="nan": continue
            p = min(max(safe_num(row.get("Tamamlanma (%)",0)),0),100)
            ev = "" if pd.isna(row.get("Kanıt / Saha Notu")) else str(row.get("Kanıt / Saha Notu"))
            con.execute("""INSERT INTO progress_records(project,wbs,progress,evidence,updated_at)
                VALUES(?,?,?,?,?) ON CONFLICT(project,wbs) DO UPDATE SET
                progress=excluded.progress,evidence=excluded.evidence,updated_at=excluded.updated_at""",
                (project,wbs,p,ev,datetime.now().isoformat(timespec="seconds")))

def db_read_stock(project):
    with db_connect() as con:
        return pd.read_sql_query("SELECT * FROM stock_records WHERE project = ? ORDER BY id DESC",
                                 con, params=(project,))

def db_save_stock(project, material, cat, unit, q_in, q_out, price, supplier):
    with db_connect() as con:
        con.execute("""INSERT INTO stock_records(project, material_name, category, unit,
            quantity_in, quantity_out, unit_price, supplier, update_date)
            VALUES(?,?,?,?,?,?,?,?,?)""",
            (project, material, cat, unit, q_in, q_out, price, supplier, str(date.today())))

def db_read_cashflow(project):
    with db_connect() as con:
        df = pd.read_sql_query("SELECT ay, nakit_akisi FROM cashflow_records WHERE project = ? ORDER BY id ASC",
                               con, params=(project,))
    if df.empty:
        default_data = [
            ("2026-01", 350000.0), ("2026-02", 750000.0), ("2026-03", 950000.0),
            ("2026-04", 800000.0), ("2026-05", 600000.0), ("2026-06", 450000.0),
            ("2026-07", 300000.0), ("2026-08", 150000.0), ("2026-09", 50000.0), ("2026-10", 0.0)
        ]
        df = pd.DataFrame(default_data, columns=["ay", "nakit_akisi"])
    return df

def db_save_cashflow(project, df):
    with db_connect() as con:
        con.execute("DELETE FROM cashflow_records WHERE project = ?", (project,))
        for _, row in df.iterrows():
            con.execute("""INSERT INTO cashflow_records(project, ay, nakit_akisi)
                VALUES(?,?,?)""", (project, str(row["ay"]), float(row["nakit_akisi"])))

def db_save_snapshot(project, payload: dict):
    with db_connect() as con:
        con.execute("""INSERT INTO project_snapshots(project,payload,updated_at)
            VALUES(?,?,?) ON CONFLICT(project) DO UPDATE SET
            payload=excluded.payload, updated_at=excluded.updated_at""",
            (project, json.dumps(payload, default=str, ensure_ascii=False).encode("utf-8"),
             datetime.now().isoformat(timespec="seconds")))

def db_load_snapshot(project):
    with db_connect() as con:
        row = con.execute("SELECT payload FROM project_snapshots WHERE project = ?",
                          (project,)).fetchone()
    if not row: return None
    try: return json.loads(row["payload"].decode("utf-8"))
    except Exception: return None

def audit_log(project, tablo, kayit_id, alan, eski, yeni):
    try:
        with db_connect() as con:
            con.execute("""INSERT INTO audit_log(project,tablo,kayit_id,alan,eski,yeni)
                VALUES(?,?,?,?,?,?)""", (project,tablo,str(kayit_id),alan,str(eski),str(yeni)))
    except Exception:
        pass

def audit_read(project, limit=200):
    with db_connect() as con:
        return pd.read_sql_query("""SELECT * FROM audit_log WHERE project = ?
            ORDER BY id DESC LIMIT ?""", con, params=(project, limit))

try:
    init_local_db()
except Exception as e:
    st.warning(f"Yerel DB başlatılamadı: {e}")

# ============================================================
# EXCEL İÇE AKTARMA
# ============================================================
def read_uploaded_excel(upload):
    return pd.read_excel(upload, sheet_name=None)

def map_imported_sheet(df, kind):
    df = df.copy()
    if kind=="plan":
        mapping = {"WBS":["WBS","WBS Kodu","Aktivite Kodu"],
            "Poz No":["Poz No","Poz","İmalat Kodu"],
            "İş Kalemi":["İş Kalemi","Açıklama","İmalat","Kalem"],
            "Ana Kategori":["Ana Kategori","Kategori","İş Grubu"],
            "Alt Kategori":["Alt Kategori","Alt Grup"],
            "Birim":["Birim","Ölçü Birimi"],
            "Miktar":["Miktar","Metraj","Adet","Miktar (m²)"],
            "Birim Fiyat":["Birim Fiyat","BF","Fiyat"],
            "Başlangıç":["Başlangıç","Başlangıç Tarihi"],
            "Bitiş":["Bitiş","Bitiş Tarihi"],
            "Sorumlu":["Sorumlu","Ekip","Firma"],
            "Açıklama":["Açıklama","Not"]}
        out = pd.DataFrame()
        for target, cands in mapping.items():
            col = find_column(df, cands); out[target] = df[col] if col else np.nan
        out["Ana Kategori"] = out["Ana Kategori"].fillna("Diğer")
        out["Miktar"] = out["Miktar"].apply(safe_num)
        out["Birim Fiyat"] = out["Birim Fiyat"].apply(safe_num)
        out = calc_plan_amount(out)
        return normalize_dates(out, ["Başlangıç","Bitiş"])[PLAN_COLS]
    if kind=="actual":
        mapping = {"WBS":["WBS","WBS Kodu","Aktivite Kodu"],
            "Poz No":["Poz No","Poz"],
            "İş Kalemi":["İş Kalemi","Açıklama","İmalat"],
            "Ana Kategori":["Ana Kategori","Kategori","İş Grubu"],
            "Birim":["Birim"],
            "Gerçekleşen Miktar":["Gerçekleşen Miktar","Miktar","Metraj"],
            "Gerçekleşen Birim Fiyat":["Gerçekleşen Birim Fiyat","Birim Fiyat","BF"],
            "Tarih":["Tarih","Fatura Tarihi","İşlem Tarihi"],
            "Tedarikçi":["Tedarikçi","Firma"],
            "Belge No":["Belge No","Fatura No","Evrak No"],
            "Açıklama":["Açıklama","Not"]}
        out = pd.DataFrame()
        for target, cands in mapping.items():
            col = find_column(df, cands); out[target] = df[col] if col else np.nan
        out["Gerçekleşen Miktar"] = out["Gerçekleşen Miktar"].apply(safe_num)
        out["Gerçekleşen Birim Fiyat"] = out["Gerçekleşen Birim Fiyat"].apply(safe_num)
        out = calc_actual_amount(out)
        return normalize_dates(out, ["Tarih"])[ACTUAL_COLS]
    if kind=="schedule":
        mapping = {"WBS":["WBS","WBS Kodu","Aktivite Kodu"],
            "Aktivite":["Aktivite","İş Kalemi","Açıklama","Görev"],
            "Süre (gün)":["Süre (gün)","Süre","Duration"],
            "Öncül WBS":["Öncül WBS","Öncül","Predecessor","Önceki Aktivite"],
            "Bağımlılık Tipi":["Bağımlılık Tipi","İlişki Tipi","Tip"],
            "Başlangıç":["Başlangıç","Başlangıç Tarihi","Plan Başlangıç"],
            "Bitiş":["Bitiş","Bitiş Tarihi","Plan Bitiş"],
            "Fiili Başlangıç":["Fiili Başlangıç","Gerçek Başlangıç"],
            "Fiili Bitiş":["Fiili Bitiş","Gerçek Bitiş"],
            "Tamamlanma (%)":["Tamamlanma (%)","İlerleme","İlerleme (%)","Fiili İlerleme %"],
            "Ağırlık (%)":["Ağırlık (%)","Ağırlık","Weight"],
            "Sorumlu":["Sorumlu","Ekip","Firma"],
            "Durum":["Durum","Status"],
            "Kanıt / Saha Notu":["Kanıt / Saha Notu","Saha Notu","Kanıt"]}
        out = pd.DataFrame()
        for target, cands in mapping.items():
            col = find_column(df, cands); out[target] = df[col] if col else np.nan
        out["Süre (gün)"] = out["Süre (gün)"].apply(safe_num)
        out["Bağımlılık Tipi"] = out["Bağımlılık Tipi"].fillna("FS").astype(str).str.upper()
        out["Tamamlanma (%)"] = out["Tamamlanma (%)"].apply(safe_num).clip(0,100)
        out["Ağırlık (%)"] = out["Ağırlık (%)"].apply(safe_num)
        out = normalize_dates(out, ["Başlangıç","Bitiş","Fiili Başlangıç","Fiili Bitiş"])
        return out[SCHEDULE_COLS]
    return df

def demo_data():
    plan = pd.DataFrame([
        ["1.1","01.001","Mimari ve uygulama projeleri","Mimarlık / Proje Gideri","Proje","LS",1,850000,850000,date(2026,1,5),date(2026,2,15),"Proje",""],
        ["1.2","02.001","Yapı denetim hizmeti","Yapı Denetim Gideri","Denetim","LS",1,420000,420000,date(2026,1,5),date(2026,12,30),"Yapı Denetim",""],
        ["1.3","03.001","Zemin ve beton laboratuvar deneyleri","Yapı Analiz Laboratuvar Gideri","Laboratuvar","LS",1,180000,180000,date(2026,1,10),date(2026,12,15),"Laboratuvar",""],
        ["2.1","04.001","Temel betonarme imalatı","Grobeton ve Beton","Beton","m³",450,5200,2340000,date(2026,2,1),date(2026,3,15),"Kaba Yapı",""],
        ["2.2","04.002","Kalıp ve demir toplam işçiliği","Toplam Kalıp + Demir İşçiliği","İşçilik","m²",3200,950,3040000,date(2026,2,1),date(2026,5,30),"Kaba Yapı",""],
        ["3.1","05.001","Elektrik ve temiz/pis su işçiliği","Elektrik + Su İşçiliği","İşçilik","LS",1,1250000,1250000,date(2026,5,1),date(2026,10,30),"MEP",""],
    ], columns=PLAN_COLS)
    actual = pd.DataFrame([
        ["1.1","01.001","Mimari ve uygulama projeleri","Mimarlık / Proje Gideri","LS",1,800000,800000,date(2026,2,1),"Proje Ofisi","F-001",""],
        ["1.2","02.001","Yapı denetim hizmeti","Yapı Denetim Gideri","LS",0.5,420000,210000,date(2026,3,1),"Denetim AŞ","F-002",""],
        ["2.1","04.001","Temel betonarme imalatı","Grobeton ve Beton","m³",180,5350,963000,date(2026,3,10),"Beton AŞ","F-003",""],
        ["2.2","04.002","Kalıp ve demir toplam işçiliği","Toplam Kalıp + Demir İşçiliği","m²",900,950,855000,date(2026,4,1),"Kaba Yapı Ltd","H-001",""],
    ], columns=ACTUAL_COLS)
    schedule = pd.DataFrame([
        ["1.1","Proje ve ruhsat",40,"",date(2026,1,5),date(2026,2,15),100,"Proje","Tamamlandı"],
        ["1.2","Şantiye kurulumu",15,"1.1",date(2026,2,16),date(2026,3,2),100,"Şantiye","Tamamlandı"],
        ["2.1","Temel imalatları",45,"1.2",date(2026,3,3),date(2026,4,17),65,"Kaba Yapı","Devam ediyor"],
        ["2.2","Taşıyıcı sistem",90,"2.1",date(2026,4,18),date(2026,7,17),15,"Kaba Yapı","Devam ediyor"],
        ["3.1","Elektrik ve mekanik",100,"2.2",date(2026,7,18),date(2026,10,25),0,"MEP","Başlamadı"],
    ], columns=["WBS","Aktivite","Süre (gün)","Öncül WBS","Başlangıç","Bitiş","Tamamlanma (%)","Sorumlu","Durum"])
    schedule["Bağımlılık Tipi"] = "FS"
    schedule["Ağırlık (%)"] = schedule["Süre (gün)"].astype(float)
    schedule["Fiili Başlangıç"] = pd.NaT
    schedule["Fiili Bitiş"] = pd.NaT
    schedule["Kanıt / Saha Notu"] = ""
    schedule = schedule[SCHEDULE_COLS]
    risk = pd.DataFrame([
        ["R-01","Beton ve demir fiyat artışı","Maliyet",4,5,20,"Alternatif tedarikçi ve erken satın alma","Satın Alma","Açık",650000],
        ["R-02","Kritik malzeme termin gecikmesi","Tedarik",3,4,12,"Onaylı tedarikçi ve termin takibi","Satın Alma","Açık",250000],
    ], columns=RISK_COLS)
    return plan, actual, schedule, risk

# ============================================================
# OTURUM
# ============================================================
if "plan" not in st.session_state:
    st.session_state.plan, st.session_state.actual, st.session_state.schedule, st.session_state.risk = demo_data()
if "project_name" not in st.session_state:
    st.session_state.project_name = "Emir Asel Konut Projesi"

# ============================================================
# BAŞLIK
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
    for logo_c in ["logo.png","logo.webp","logo.jpg","logo.jpeg"]:
        lp = Path(__file__).with_name(logo_c)
        if lp.exists():
            st.image(str(lp), width=190); break
    st.markdown("## 🏗️ EMİR ASEL")
    st.caption(f"Proje Kontrol Sistemi | V{APP_VERSION}")
    MENU = ["Yönetici Dashboard","Proje Künyesi","Bütçe ve Metraj",
            "Gerçekleşen Maliyet","İş Programı / CPM","EVM Performans",
            "Satın Alma ve Tedarik","İşçilik ve Puantaj","Nakit Akışı (S-Curve)",
            "Şantiye Stok & Depo","Risk Yönetimi","Veri Kontrolü",
            "📄 Raporlar (PDF/JSON)","🕓 Değişiklik Geçmişi","⚙️ Sistem ve Yedek"]
    page = st.radio("MODÜLLER", MENU)
    st.divider()
    project_name = st.text_input("Proje Adı", st.session_state.project_name)
    if project_name != st.session_state.project_name:
        st.session_state.project_name = project_name
    company = st.text_input("Firma", "Emir Asel İnşaat")
    currency = st.selectbox("Para Birimi", ["₺","$","€"], index=0)
    vat_rate = st.number_input("KDV oranı", min_value=0.0, max_value=1.0,
                                value=VAT_DEFAULT, step=0.01, format="%.2f")
    st.markdown("---")
    try:
        with db_connect() as _c:
            _n = _c.execute("SELECT COUNT(*) AS n FROM progress_records").fetchone()["n"]
        st.markdown(f'<span class="badge-db">💾 SQLite · {_n} ilerleme kaydı</span>',
                    unsafe_allow_html=True)
    except Exception:
        st.markdown('<span class="badge-warn">⚠️ DB bağlantı sorunu</span>',
                    unsafe_allow_html=True)
    st.markdown("---")
    if st.button("🔄 Önbelleği Temizle"):
        st.cache_data.clear(); st.rerun()

# ============================================================
# VERİ HAZIRLA
# ============================================================
plan = st.session_state.plan.copy()
actual = st.session_state.actual.copy()
schedule = st.session_state.schedule.copy()
risk = st.session_state.risk.copy()

try:
    saved = db_load_progress(project_name)
    for i, row in schedule.iterrows():
        s = saved.get(str(row.get("WBS","")).strip())
        if s:
            schedule.at[i,"Tamamlanma (%)"] = s["progress"]
            if s.get("evidence"): schedule.at[i,"Kanıt / Saha Notu"] = s["evidence"]
except Exception:
    pass

plan = calc_plan_amount(ensure_columns(plan, PLAN_COLS))
actual = calc_actual_amount(ensure_columns(actual, ACTUAL_COLS))
budget = float(plan["Planlanan Tutar"].sum())
actual_cost = float(actual["Gerçekleşen Tutar"].sum())
cost_summary = summarize_costs(plan, actual)
progress = weighted_progress(schedule)
evm = calculate_evm(plan, actual, schedule)
cpm, cpm_warnings = calculate_cpm(schedule)

# ============================================================
# SAYFA: DASHBOARD
# ============================================================
if page == "Yönetici Dashboard":
    st.title("Yönetici Dashboard")
    st.caption(f"{company} | {project_name} | Son güncelleme: {datetime.now():%d.%m.%Y %H:%M}")
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Toplam Bütçe", money(budget,currency))
    c2.metric("Gerçekleşen Maliyet", money(actual_cost,currency),
              delta=money(actual_cost-budget,currency) if budget else None,
              delta_color="inverse")
    c3.metric("Fiziksel İlerleme", pct(progress))
    c4.metric("EAC", money(evm["EAC"],currency) if pd.notna(evm["EAC"]) else "Hesaplanamadı")

    st.markdown('<div class="section-title">📊 Kategori Bazında Bütçe / Gerçekleşen</div>',
                unsafe_allow_html=True)
    left, right = st.columns([1.3,1])
    with left:
        chart_df = cost_summary.melt(id_vars="Ana Kategori",
                                     value_vars=["Bütçe","Gerçekleşen"],
                                     var_name="Tür", value_name="Tutar")
        if not chart_df.empty:
            fig = px.bar(chart_df, x="Ana Kategori", y="Tutar", color="Tür",
                         barmode="group", color_discrete_sequence=[NAVY, RED])
            fig.update_layout(height=440, xaxis_tickangle=-35,
                              margin=dict(l=10,r=10,t=20,b=120),
                              plot_bgcolor="white", paper_bgcolor="white")
            st.plotly_chart(fig, use_container_width=True)
    with right:
        pie = cost_summary[cost_summary["Bütçe"]>0]
        if not pie.empty:
            fig = px.pie(pie, names="Ana Kategori", values="Bütçe",
                         hole=.55, color_discrete_sequence=PALETTE)
            fig.update_layout(height=440, margin=dict(l=5,r=5,t=20,b=5))
            st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="section-title">🎯 Yönetim Göstergeleri</div>',
                unsafe_allow_html=True)
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("CPI", f'{evm["CPI"]:.2f}' if pd.notna(evm["CPI"]) else "—")
    m2.metric("SPI", f'{evm["SPI"]:.2f}' if pd.notna(evm["SPI"]) else "—")
    m3.metric("Açık Risk",
              int((risk["Durum"].astype(str).str.lower()=="açık").sum()) if not risk.empty else 0)
    m4.metric("Kritik Aktivite",
              int((cpm["Kritik mi?"]=="EVET").sum()) if not cpm.empty and "Kritik mi?" in cpm else 0)
    st.info("CPI/SPI göstergeleri WBS eşleştirmesi ve girilen ilerleme verilerine bağlıdır.")

# ============================================================
# SAYFA: PROJE KÜNYESİ
# ============================================================
elif page == "Proje Künyesi":
    st.title("Proje Künyesi ve Veri Yükleme")
    uploaded = st.file_uploader("Excel dosyası yükle (.xlsx)", type=["xlsx"])
    if uploaded:
        try:
            sheets = read_uploaded_excel(uploaded)
            st.write("Bulunan sayfalar:", list(sheets.keys()))
            plan_sheet = st.selectbox("Bütçe/metraj sayfası", list(sheets.keys()), key="pl_sheet")
            actual_sheet = st.selectbox("Gerçekleşen sayfa", ["—"]+list(sheets.keys()), key="ac_sheet")
            schedule_sheet = st.selectbox("İş programı sayfası", ["—"]+list(sheets.keys()), key="sc_sheet")
            if st.button("Seçilen sayfaları içe aktar", type="primary"):
                st.session_state.plan = map_imported_sheet(sheets[plan_sheet], "plan")
                if actual_sheet != "—":
                    st.session_state.actual = map_imported_sheet(sheets[actual_sheet], "actual")
                if schedule_sheet != "—":
                    st.session_state.schedule = map_imported_sheet(sheets[schedule_sheet], "schedule")
                audit_log(project_name, "excel", uploaded.name, "import", "-", "ok")
                st.success("İçe aktarma tamamlandı."); st.rerun()
        except Exception as e:
            st.error(f"Excel okunamadı: {e}")

    st.markdown('<div class="section-title">💾 Proje Snapshot</div>', unsafe_allow_html=True)
    colA, colB = st.columns(2)
    with colA:
        if st.button("💾 Mevcut projeyi snapshot olarak kaydet"):
            payload = {
                "plan": plan.to_dict(orient="records"),
                "actual": actual.to_dict(orient="records"),
                "schedule": schedule.to_dict(orient="records"),
                "risk": risk.to_dict(orient="records"),
                "saved_at": datetime.now().isoformat(timespec="seconds"),
            }
            db_save_snapshot(project_name, payload)
            st.success("✅ Snapshot kaydedildi.")
    with colB:
        if st.button("📂 Snapshot'tan geri yükle"):
            snap = db_load_snapshot(project_name)
            if snap:
                st.session_state.plan = pd.DataFrame(snap["plan"], columns=PLAN_COLS)
                st.session_state.actual = pd.DataFrame(snap["actual"], columns=ACTUAL_COLS)
                st.session_state.schedule = pd.DataFrame(snap["schedule"], columns=SCHEDULE_COLS)
                st.session_state.risk = pd.DataFrame(snap["risk"], columns=RISK_COLS)
                st.success("✅ Geri yüklendi."); st.rerun()
            else:
                st.warning("Kayıtlı snapshot yok.")

    st.markdown('<div class="section-title">📄 Veri Şablonları</div>', unsafe_allow_html=True)
    st.download_button("Boş veri şablonlarını indir",
        excel_export({
            "Butce_Metraj": pd.DataFrame(columns=PLAN_COLS),
            "Gerceklesen": pd.DataFrame(columns=ACTUAL_COLS),
            "Is_Programi": pd.DataFrame(columns=SCHEDULE_COLS),
            "Riskler": pd.DataFrame(columns=RISK_COLS),
        }),
        file_name="Emir_Asel_V82_Veri_Sablonlari.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============================================================
# SAYFA: BÜTÇE VE METRAJ
# ============================================================
elif page == "Bütçe ve Metraj":
    st.title("Bütçe ve Metraj Yönetimi")
    st.caption("Planlanan tutar = Miktar × Birim Fiyat.")
    edited = st.data_editor(plan, num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "Ana Kategori": st.column_config.SelectboxColumn(options=COST_CATEGORIES),
            "Miktar": st.column_config.NumberColumn(min_value=0, format="%.3f"),
            "Birim Fiyat": st.column_config.NumberColumn(min_value=0, format="%.2f"),
            "Planlanan Tutar": st.column_config.NumberColumn(disabled=True, format="%.2f"),
            "Başlangıç": st.column_config.DateColumn(),
            "Bitiş": st.column_config.DateColumn(),
        }, key="plan_editor")
    if st.button("Bütçe / metraj değişikliklerini kaydet", type="primary"):
        edited = ensure_columns(edited, PLAN_COLS)
        edited["Miktar"] = edited["Miktar"].apply(safe_num)
        edited["Birim Fiyat"] = edited["Birim Fiyat"].apply(safe_num)
        st.session_state.plan = calc_plan_amount(edited)
        audit_log(project_name, "plan", "bulk", "update", "-", f"{len(edited)} satır")
        st.success("Kaydedildi."); st.rerun()
    st.markdown('<div class="section-title">📊 Kategori Özeti</div>', unsafe_allow_html=True)
    st.dataframe(cost_summary, use_container_width=True, hide_index=True)
    st.metric("Toplam Planlanan Bütçe", money(budget, currency))

# ============================================================
# SAYFA: GERÇEKLEŞEN
# ============================================================
elif page == "Gerçekleşen Maliyet":
    st.title("Gerçekleşen Maliyet / Satın Alma Kayıtları")
    edited = st.data_editor(actual, num_rows="dynamic", use_container_width=True, hide_index=True,
        column_config={
            "Ana Kategori": st.column_config.SelectboxColumn(options=COST_CATEGORIES),
            "Gerçekleşen Miktar": st.column_config.NumberColumn(min_value=0, format="%.3f"),
            "Gerçekleşen Birim Fiyat": st.column_config.NumberColumn(min_value=0, format="%.2f"),
            "Gerçekleşen Tutar": st.column_config.NumberColumn(disabled=True, format="%.2f"),
            "Tarih": st.column_config.DateColumn(),
        }, key="actual_editor")
    if st.button("Gerçekleşen kayıtları kaydet", type="primary"):
        edited = ensure_columns(edited, ACTUAL_COLS)
        edited["Gerçekleşen Miktar"] = edited["Gerçekleşen Miktar"].apply(safe_num)
        edited["Gerçekleşen Birim Fiyat"] = edited["Gerçekleşen Birim Fiyat"].apply(safe_num)
        st.session_state.actual = calc_actual_amount(edited)
        audit_log(project_name, "actual", "bulk", "update", "-", f"{len(edited)} satır")
        st.success("Kaydedildi."); st.rerun()
    st.markdown('<div class="section-title">📊 Kategori Bazında Sapma</div>', unsafe_allow_html=True)
    st.dataframe(cost_summary, use_container_width=True, hide_index=True)
    over = cost_summary[cost_summary["Sapma"]>0]
    if not over.empty:
        st.warning("Bütçeyi aşan kategoriler: " + ", ".join(over["Ana Kategori"].astype(str).tolist()))

# ============================================================
# SAYFA: İŞ PROGRAMI / CPM
# ============================================================
elif page == "İş Programı / CPM":
    st.title("İş Programı, Gantt ve Kritik Yol")
    st.caption("Öncül WBS'leri virgül/noktalı virgülle ayırın. Farklı tip için `A01:FS; A02:SS`.")
    edited = st.data_editor(schedule, num_rows="dynamic", use_container_width=True,
        hide_index=True,
        column_config={
            "Süre (gün)": st.column_config.NumberColumn(min_value=0, step=1),
            "Bağımlılık Tipi": st.column_config.SelectboxColumn(options=["FS","SS","FF","SF"]),
            "Tamamlanma (%)": st.column_config.NumberColumn(min_value=0, max_value=100, format="%.1f"),
            "Ağırlık (%)": st.column_config.NumberColumn(min_value=0, format="%.2f"),
            "Başlangıç": st.column_config.DateColumn(),
            "Bitiş": st.column_config.DateColumn(),
            "Fiili Başlangıç": st.column_config.DateColumn(),
            "Fiili Bitiş": st.column_config.DateColumn(),
        }, key="schedule_editor")
    if st.button("İş programını kaydet", type="primary"):
        edited = ensure_columns(edited, SCHEDULE_COLS)
        edited["Süre (gün)"] = edited["Süre (gün)"].apply(safe_num)
        edited["Tamamlanma (%)"] = edited["Tamamlanma (%)"].apply(safe_num).clip(0,100)
        edited["Ağırlık (%)"] = edited["Ağırlık (%)"].apply(safe_num)
        edited["Bağımlılık Tipi"] = edited["Bağımlılık Tipi"].fillna("FS").astype(str).str.upper()
        st.session_state.schedule = edited
        try:
            db_save_progress(project_name, edited)
            audit_log(project_name, "schedule", "bulk", "update", "-", f"{len(edited)} satır")
        except Exception as e:
            st.warning(f"DB'ye yazılamadı: {e}")
        st.success("Kaydedildi."); st.rerun()

    for w in cpm_warnings:
        st.warning(w)

    if not cpm.empty:
        st.markdown('<div class="section-title">📐 CPM Sonuçları</div>',
                    unsafe_allow_html=True)
        st.metric("Hesaplanan Proje Süresi", f'{cpm["Erken Bitiş"].max():.0f} gün')
        st.dataframe(cpm, use_container_width=True, hide_index=True)
        fig = px.bar(cpm, x="Aktivite", y="Süre (gün)", color="Kritik mi?",
                     color_discrete_map={"EVET":RED,"Hayır":TEAL})
        fig.update_layout(height=400, xaxis_tickangle=-25,
                          plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
        dated = schedule.dropna(subset=["Başlangıç","Bitiş"]).copy()
        if not dated.empty:
            fig = px.timeline(dated, x_start="Başlangıç", x_end="Bitiş",
                              y="Aktivite", color="Tamamlanma (%)",
                              color_continuous_scale="Blues")
            fig.update_yaxes(autorange="reversed")
            fig.update_layout(height=450, plot_bgcolor="white", paper_bgcolor="white")
            st.plotly_chart(fig, use_container_width=True)

# ============================================================
# SAYFA: EVM PERFORMANS
# ============================================================
elif page == "EVM Performans":
    st.title("Kazanılmış Değer Yönetimi (EVM)")
    a,b,c = st.columns(3)
    a.metric("BAC", money(evm["BAC"],currency))
    b.metric("PV", money(evm["PV"],currency))
    c.metric("EV", money(evm["EV"],currency))
    d,e,f = st.columns(3)
    d.metric("AC", money(evm["AC"],currency))
    e.metric("CPI", f'{evm["CPI"]:.3f}' if pd.notna(evm["CPI"]) else "—")
    f.metric("SPI", f'{evm["SPI"]:.3f}' if pd.notna(evm["SPI"]) else "—")
    g,h,i = st.columns(3)
    g.metric("CV", money(evm["CV"],currency))
    h.metric("SV", money(evm["SV"],currency))
    i.metric("EAC", money(evm["EAC"],currency) if pd.notna(evm["EAC"]) else "—")
    if pd.notna(evm["CPI"]):
        if evm["CPI"]<1: st.error("⚠️ Maliyet performansı bütçenin gerisinde (CPI < 1).")
        else: st.success("✅ Maliyet performansı bütçe hedefinde (CPI ≥ 1).")
    if pd.notna(evm["SPI"]):
        if evm["SPI"]<1: st.warning("⚠️ Zaman performansı planın gerisinde (SPI < 1).")
        else: st.success("✅ Zaman performansı plan hedefinde (SPI ≥ 1).")
    if not cost_summary.empty:
        melt = cost_summary.melt(id_vars="Ana Kategori",
                                 value_vars=["Bütçe","Gerçekleşen"],
                                 var_name="Tür", value_name="Tutar")
        fig = px.bar(melt, x="Ana Kategori", y="Tutar", color="Tür",
                     barmode="group", color_discrete_sequence=[NAVY, RED])
        fig.update_layout(height=400, xaxis_tickangle=-35,
                          plot_bgcolor="white", paper_bgcolor="white",
                          margin=dict(l=10,r=10,t=20,b=120))
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# SAYFA: SATIN ALMA
# ============================================================
elif page == "Satın Alma ve Tedarik":
    st.title("Satın Alma ve Tedarik Takibi")
    purchase_cols = ["order_no","item","supplier","quantity","unit","unit_price",
                     "total","order_date","due_date","status","delivered_qty",
                     "invoiced_amount","paid_amount","wbs","note"]
    purchases = db_read("purchase_records", project_name)
    if purchases.empty:
        purchases = pd.DataFrame(columns=purchase_cols)
    else:
        purchases = purchases[[c for c in purchase_cols if c in purchases.columns]]
    with st.expander("➕ Yeni sipariş", expanded=False):
        with st.form("purchase_form"):
            a,b,c = st.columns(3)
            order_no = a.text_input("Sipariş No")
            item = b.text_input("Malzeme / Hizmet")
            supplier = c.text_input("Tedarikçi")
            d,e,f = st.columns(3)
            qty = d.number_input("Miktar", min_value=0.0, step=1.0)
            unit = e.text_input("Birim", "Adet")
            unit_price = f.number_input("Birim fiyat", min_value=0.0, step=100.0)
            g,h,i = st.columns(3)
            order_date = g.date_input("Sipariş tarihi")
            due_date = h.date_input("Termin")
            status = i.selectbox("Durum", ["Teklif Alındı","Sipariş Verildi",
                                            "Kısmi Teslim","Teslim Edildi","İptal"])
            j,k,l = st.columns(3)
            delivered = j.number_input("Teslim edilen", min_value=0.0, step=1.0)
            invoiced = k.number_input("Faturalanan", min_value=0.0, step=100.0)
            paid = l.number_input("Ödenen", min_value=0.0, step=100.0)
            wbs = st.text_input("İlgili WBS")
            note = st.text_input("Açıklama")
            if st.form_submit_button("Siparişi kaydet", type="primary"):
                new_row = pd.DataFrame([[order_no,item,supplier,qty,unit,unit_price,
                                          qty*unit_price,str(order_date),str(due_date),
                                          status,delivered,invoiced,paid,wbs,note]],
                                        columns=purchase_cols)
                purchases = pd.concat([purchases,new_row], ignore_index=True)
                db_save_frame("purchase_records", purchases, project_name)
                audit_log(project_name, "satin_alma", order_no, "insert", "-", item)
                st.success("Eklendi."); st.rerun()
    if not purchases.empty:
        purchases["open_commitment"] = (purchases["total"].apply(safe_num) -
                                         purchases["invoiced_amount"].apply(safe_num)).clip(lower=0)
        x,y,z = st.columns(3)
        x.metric("Toplam Sipariş", money(purchases["total"].apply(safe_num).sum(),currency))
        y.metric("Faturalanan", money(purchases["invoiced_amount"].apply(safe_num).sum(),currency))
        z.metric("Açık Taahhüt", money(purchases["open_commitment"].sum(),currency))
        edited_p = st.data_editor(purchases, num_rows="dynamic",
                                   use_container_width=True, hide_index=True,
                                   key="purchase_editor")
        if st.button("Kayıtları güncelle"):
            edited_p["total"] = (edited_p["quantity"].apply(safe_num) *
                                 edited_p["unit_price"].apply(safe_num))
            db_save_frame("purchase_records", edited_p, project_name)
            st.success("Kaydedildi."); st.rerun()
    else:
        st.info("Henüz kayıt yok.")

# ============================================================
# SAYFA: İŞÇİLİK
# ============================================================
elif page == "İşçilik ve Puantaj":
    st.title("İşçilik ve Puantaj Takibi")
    labor_cols = ["work_date","team","work_description","worker_count",
                  "daily_wage","day_count","total","wbs","subcontractor","note"]
    labor = db_read("labor_records", project_name)
    if labor.empty:
        labor = pd.DataFrame(columns=labor_cols)
    else:
        labor = labor[[c for c in labor_cols if c in labor.columns]]
    with st.expander("➕ Yeni puantaj", expanded=False):
        with st.form("labor_form"):
            a,b,c = st.columns(3)
            work_date = a.date_input("Tarih")
            team = b.text_input("Ekip")
            description = c.text_input("İş tanımı")
            d,e,f = st.columns(3)
            workers = d.number_input("Çalışan", min_value=0.0, step=1.0)
            wage = e.number_input("Yevmiye", min_value=0.0, step=100.0)
            days = f.number_input("Gün", min_value=0.0, step=0.5)
            wbs = st.text_input("İlgili WBS")
            subcontractor = st.text_input("Taşeron")
            note = st.text_input("Not")
            if st.form_submit_button("Kaydet", type="primary"):
                new_row = pd.DataFrame([[str(work_date),team,description,workers,wage,
                                          days,workers*wage*days,wbs,subcontractor,note]],
                                        columns=labor_cols)
                labor = pd.concat([labor,new_row], ignore_index=True)
                db_save_frame("labor_records", labor, project_name)
                audit_log(project_name, "iscilik", team, "insert", "-", description)
                st.success("Eklendi."); st.rerun()
    if not labor.empty:
        labor["total"] = labor["total"].apply(safe_num)
        a,b = st.columns(2)
        a.metric("Toplam İşçilik", money(labor["total"].sum(),currency))
        b.metric("Adam-Gün", f'{(labor["worker_count"].apply(safe_num)*labor["day_count"].apply(safe_num)).sum():,.1f}')
        edited_l = st.data_editor(labor, num_rows="dynamic",
                                   use_container_width=True, hide_index=True,
                                   key="labor_editor")
        if st.button("Kayıtları güncelle"):
            edited_l["total"] = (edited_l["worker_count"].apply(safe_num) *
                                 edited_l["daily_wage"].apply(safe_num) *
                                 edited_l["day_count"].apply(safe_num))
            db_save_frame("labor_records", edited_l, project_name)
            st.success("Kaydedildi."); st.rerun()
    else:
        st.info("Henüz kayıt yok.")

# ============================================================
# SAYFA: NAKİT AKIŞI
# ============================================================
elif page == "Nakit Akışı (S-Curve)":
    st.title("Nakit Akışı ve S-Curve")
    raw_cf = db_read_cashflow(project_name)
    edited_cf = st.data_editor(raw_cf, num_rows="dynamic", use_container_width=True,
                                key="cashflow_editor")
    if st.button("💾 Değişiklikleri Kaydet", use_container_width=True):
        db_save_cashflow(project_name, edited_cf)
        st.success("Kaydedildi.")
        raw_cf = edited_cf

    scurve_df = raw_cf.copy()
    scurve_df["nakit_akisi"] = pd.to_numeric(scurve_df["nakit_akisi"], errors="coerce").fillna(0)
    scurve_df["Kümülatif"] = scurve_df["nakit_akisi"].cumsum()
    scurve_df.columns = ["Dönem / Ay", "Aylık Nakit Akışı (TL)", "Kümülatif S-Curve (TL)"]

    excel_bytes = excel_export({"S-Curve Nakit Akışı": scurve_df})
    st.download_button("📊 Excel İndir", data=excel_bytes,
        file_name=f"{project_name}_SCurve.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    fig = go.Figure()
    fig.add_trace(go.Bar(x=scurve_df["Dönem / Ay"],
                         y=scurve_df["Aylık Nakit Akışı (TL)"],
                         name="Aylık Nakit Akışı", marker_color=NAVY))
    fig.add_trace(go.Scatter(x=scurve_df["Dönem / Ay"],
                              y=scurve_df["Kümülatif S-Curve (TL)"],
                              name="Kümülatif S-Curve",
                              yaxis="y2",
                              line=dict(color=RED_ACCENT, width=4)))
    fig.update_layout(
        title="Proje Nakit Akışı ve S-Curve",
        xaxis=dict(title="Aylar"),
        yaxis=dict(title="Aylık Tutar (TL)", side="left"),
        yaxis2=dict(title="Kümülatif Tutar (TL)", overlaying="y", side="right", showgrid=False),
        legend=dict(x=0.1, y=1.1, orientation="h"),
        plot_bgcolor="white", paper_bgcolor="white", height=450,
    )
    st.plotly_chart(fig, use_container_width=True)

# ============================================================
# SAYFA: STOK & DEPO
# ============================================================
elif page == "Şantiye Stok & Depo":
    st.title("Şantiye Malzeme Stok ve Depo Takibi")
    with st.form("stock_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            mat_name = st.text_input("Malzeme Adı", placeholder="Örn: C30 Hazır Beton")
            mat_cat = st.selectbox("Kategori", ["Demir / Çelik", "Beton / Agrega",
                                                "Tuğla / Ytong", "İnce Yapı", "Diğer"])
        with col2:
            unit = st.selectbox("Birim", ["m³","Ton","Adet","Torba","mt"])
            q_in = st.number_input("Giren Miktar", min_value=0.0, value=50.0)
        with col3:
            q_out = st.number_input("Çıkan Miktar", min_value=0.0, value=10.0)
            price = st.number_input("Birim Fiyat (TL)", min_value=0.0, value=3200.0)
            supplier = st.text_input("Tedarikçi", placeholder="Beton AŞ")
        if st.form_submit_button("Stok Kaydet", type="primary"):
            db_save_stock(project_name, mat_name, mat_cat, unit, q_in, q_out, price, supplier)
            audit_log(project_name, "stok", mat_name, "insert", "-", f"{q_in} {unit}")
            st.success("Stok hareketi kaydedildi!"); st.rerun()
    stock_df = db_read_stock(project_name)
    if not stock_df.empty:
        stock_df["Kalan Stok"] = stock_df["quantity_in"] - stock_df["quantity_out"]
        c1, c2 = st.columns(2)
        c1.metric("Toplam Giren", f'{stock_df["quantity_in"].sum():,.2f}')
        c2.metric("Toplam Çıkan", f'{stock_df["quantity_out"].sum():,.2f}')
        st.dataframe(stock_df, use_container_width=True)
    else:
        st.info("Henüz stok hareketi yok.")

# ============================================================
# SAYFA: RİSK
# ============================================================
elif page == "Risk Yönetimi":
    st.title("Risk Yönetimi")
    edited = st.data_editor(risk, num_rows="dynamic", use_container_width=True,
        hide_index=True,
        column_config={
            "Olasılık (1-5)": st.column_config.NumberColumn(min_value=1, max_value=5, step=1),
            "Etki (1-5)": st.column_config.NumberColumn(min_value=1, max_value=5, step=1),
            "Tahmini Maliyet Etkisi": st.column_config.NumberColumn(min_value=0, format="%.2f"),
            "Durum": st.column_config.SelectboxColumn(options=["Açık","İzleniyor","Kapalı"]),
        }, key="risk_editor")
    if st.button("Riskleri kaydet", type="primary"):
        edited = ensure_columns(edited, RISK_COLS)
        edited["Olasılık (1-5)"] = edited["Olasılık (1-5)"].apply(safe_num).clip(1,5)
        edited["Etki (1-5)"] = edited["Etki (1-5)"].apply(safe_num).clip(1,5)
        edited["Risk Skoru"] = edited["Olasılık (1-5)"] * edited["Etki (1-5)"]
        st.session_state.risk = edited
        audit_log(project_name, "risk", "bulk", "update", "-", f"{len(edited)} satır")
        st.success("Kaydedildi."); st.rerun()
    if not risk.empty:
        risk["Risk Skoru"] = (risk["Olasılık (1-5)"].apply(safe_num) *
                                risk["Etki (1-5)"].apply(safe_num))
        fig = px.scatter(risk, x="Olasılık (1-5)", y="Etki (1-5)",
                         size="Tahmini Maliyet Etkisi", color="Risk Skoru",
                         hover_name="Risk Tanımı", text="Risk ID",
                         color_continuous_scale="RdYlGn_r",
                         range_x=[0.5,5.5], range_y=[0.5,5.5])
        fig.update_layout(height=480, plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
        exposure = (risk["Tahmini Maliyet Etkisi"].apply(safe_num) *
                    risk["Olasılık (1-5)"].apply(safe_num) / 5).sum()
        st.metric("Olasılık Ağırlıklı Risk Maliyeti", money(exposure, currency))

# ============================================================
# SAYFA: VERİ KONTROLÜ
# ============================================================
elif page == "Veri Kontrolü":
    st.title("Veri Kalitesi Kontrolleri")
    checks = []
    def add_check(name, count, detail):
        checks.append({"Kontrol":name,"Hata / Uyarı":int(count),"Açıklama":detail})
    add_check("Boş WBS – bütçe", plan["WBS"].isna().sum(), "Bütçe satırlarında WBS bulunmalı.")
    add_check("Tekrarlı WBS – bütçe", plan["WBS"].duplicated().sum(), "Aynı WBS birden fazla pozda.")
    add_check("Negatif metraj", (plan["Miktar"].apply(safe_num)<0).sum(), "Miktar negatif olmamalı.")
    add_check("Negatif birim fiyat", (plan["Birim Fiyat"].apply(safe_num)<0).sum(), "Fiyat negatif olmamalı.")
    add_check("Bitiş < Başlangıç",
              (pd.to_datetime(plan["Bitiş"],errors="coerce") <
               pd.to_datetime(plan["Başlangıç"],errors="coerce")).sum(),
              "Tarih aralığı ters.")
    add_check("WBS'siz gerçekleşen", actual["WBS"].isna().sum(), "Gerçekleşen bütçe WBS'ine bağlanmalı.")
    add_check("Bütçede olmayan WBS (gerçekleşen)",
              (~actual["WBS"].astype(str).isin(plan["WBS"].astype(str))).sum(),
              "Bütçede karşılığı yok.")
    add_check("İlerleme aralık dışı",
              ((schedule["Tamamlanma (%)"].apply(safe_num)<0) |
               (schedule["Tamamlanma (%)"].apply(safe_num)>100)).sum(),
              "İlerleme %0–100 aralığında olmalı.")
    check_df = pd.DataFrame(checks)
    st.dataframe(check_df, use_container_width=True, hide_index=True)
    total = check_df["Hata / Uyarı"].sum()
    if total == 0: st.success("✅ Temel kontrollerde hata bulunmadı.")
    else: st.warning(f"⚠️ {total} adet kontrol uyarısı.")
    st.caption("Bu kontroller veri giriş kalitesini değerlendirir.")

# ============================================================
# SAYFA: RAPORLAR
# ============================================================
elif page == "📄 Raporlar (PDF/JSON)":
    st.title("Raporlar & Dışa Aktarım")

    st.markdown('<div class="section-title">🏗️ Otomatik Metraj</div>', unsafe_allow_html=True)
    hcol1, hcol2 = st.columns(2)
    with hcol1:
        st.subheader("Beton Hacmi")
        b_en = st.number_input("En (m)", min_value=0.1, value=5.0, key="b_en")
        b_boy = st.number_input("Boy (m)", min_value=0.1, value=10.0, key="b_boy")
        b_yuk = st.number_input("Yükseklik (m)", min_value=0.01, value=0.3, key="b_yuk")
        adet = st.number_input("Adet", min_value=1, value=1, key="b_adet")
        st.success(f"Beton: **{b_en*b_boy*b_yuk*adet:,.2f} m³**")
    with hcol2:
        st.subheader("Donatı Metrajı")
        cap = st.selectbox("Çap", ["Ø8","Ø10","Ø12","Ø14","Ø16","Ø20","Ø25","Ø32"], key="cap")
        boy_m = st.number_input("Toplam Boy (mt)", min_value=1.0, value=1200.0, key="boy_m")
        birim_agirlik = {"Ø8":0.395,"Ø10":0.617,"Ø12":0.888,"Ø14":1.21,
                         "Ø16":1.579,"Ø20":2.466,"Ø25":3.855,"Ø32":6.313}
        kg = boy_m * birim_agirlik[cap]
        st.info(f"Donatı: **{kg:,.2f} kg ({kg/1000:,.2f} Ton)**")

    st.markdown('<div class="section-title">📥 Excel Raporu</div>', unsafe_allow_html=True)
    excel_bytes = excel_export({
        "Proje_Ozeti": pd.DataFrame([
            ["Firma",company],["Proje",project_name],["Sürüm",APP_VERSION],
            ["Tarih",datetime.now().strftime("%d.%m.%Y")],
            ["Toplam Bütçe",budget],["Gerçekleşen",actual_cost],
            ["İlerleme (%)",progress],["KDV",vat_rate]],
            columns=["Gösterge","Değer"]),
        "Butce_Metraj": plan,
        "Gerceklesen": actual,
        "Kategori_Sapma": cost_summary,
        "Is_Programi": schedule,
        "CPM": cpm,
        "Riskler": risk,
        "Satin_Alma": db_read("purchase_records",project_name),
        "Iscilik": db_read("labor_records",project_name),
    })
    st.download_button("📥 Excel Raporu İndir", excel_bytes,
        file_name=f"EmirAsel_Rapor_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown('<div class="section-title">📦 JSON Yedek</div>', unsafe_allow_html=True)
    yedek = {
        "tarih": datetime.now().isoformat(),
        "proje": project_name, "firma": company,
        "plan": plan.to_dict(orient="records"),
        "actual": actual.to_dict(orient="records"),
        "schedule": schedule.to_dict(orient="records"),
        "risk": risk.to_dict(orient="records"),
        "cpm": cpm.to_dict(orient="records") if not cpm.empty else [],
        "evm": {k: (float(v) if pd.notna(v) else None) for k,v in evm.items()},
    }
    st.download_button("📥 JSON Yedek İndir",
        json.dumps(yedek, ensure_ascii=False, indent=2, default=str),
        file_name=f"EmirAsel_Yedek_{datetime.now():%Y%m%d_%H%M}.json",
        mime="application/json")

    st.markdown('<div class="section-title">📄 PDF Rapor</div>', unsafe_allow_html=True)
    if not PDF_OK:
        st.warning("ReportLab yüklü değil. `pip install reportlab` çalıştırın.")
    elif st.button("📄 PDF Rapor Oluştur", type="primary"):
        buf = BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=1.5*cm, bottomMargin=1.5*cm)
        story = []
        sty = getSampleStyleSheet()
        title_st = ParagraphStyle('T', parent=sty['Title'], fontSize=18,
                                    textColor=rl_colors.HexColor(NAVY),
                                    alignment=1, spaceAfter=8)
        sub_st = ParagraphStyle('S', parent=sty['Normal'], fontSize=10,
                                 textColor=rl_colors.grey, alignment=1, spaceAfter=20)
        story.append(Paragraph("EMİR ASEL İNŞAAT", title_st))
        story.append(Paragraph(
            f"Proje Yönetim Raporu · {datetime.now():%d.%m.%Y %H:%M}", sub_st))
        ozet_data = [["Gösterge","Değer"],
                     ["Toplam Bütçe", money(budget,currency)],
                     ["Gerçekleşen", money(actual_cost,currency)],
                     ["Fiziksel İlerleme", pct(progress)],
                     ["CPI", f'{evm["CPI"]:.2f}' if pd.notna(evm["CPI"]) else "—"],
                     ["SPI", f'{evm["SPI"]:.2f}' if pd.notna(evm["SPI"]) else "—"],
                     ["Kritik Aktivite",
                      str(len(cpm[cpm["Kritik mi?"]=="EVET"])) if not cpm.empty else "0"]]
        t = Table(ozet_data, colWidths=[8*cm, 8*cm])
        t.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0), rl_colors.HexColor(NAVY)),
            ('TEXTCOLOR',(0,0),(-1,0), rl_colors.white),
            ('FONTNAME',(0,0),(-1,0), 'Helvetica-Bold'),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('GRID',(0,0),(-1,-1), 0.5, rl_colors.grey),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),
             [rl_colors.white, rl_colors.HexColor("#F8F9FA")]),
        ]))
        story.append(t)
        story.append(Spacer(1, 0.8*cm))
        story.append(Paragraph("CPM Tablosu", sty['Heading2']))
        cpm_rows = [["WBS","İş Kalemi","Süre","ES","EF","Kritik"]]
        for _, r in cpm.head(30).iterrows():
            cpm_rows.append([r["WBS"], r["Aktivite"][:30],
                             str(r["Süre (gün)"]),
                             str(r["Erken Başlangıç"]),
                             str(r["Erken Bitiş"]),
                             r["Kritik mi?"]])
        tb = Table(cpm_rows, colWidths=[2*cm, 6*cm, 2*cm, 2*cm, 2*cm, 2*cm])
        tb.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,0), rl_colors.HexColor(NAVY)),
            ('TEXTCOLOR',(0,0),(-1,0), rl_colors.white),
            ('FONTNAME',(0,0),(-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',(0,0),(-1,-1), 8),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('GRID',(0,0),(-1,-1), 0.3, rl_colors.grey),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),
             [rl_colors.white, rl_colors.HexColor("#F1F5F9")]),
        ]))
        story.append(tb)
        doc.build(story); buf.seek(0)
        st.download_button("📥 PDF İndir", data=buf,
            file_name=f"EmirAsel_Rapor_{datetime.now():%Y%m%d_%H%M}.pdf",
            mime="application/pdf")

# ============================================================
# SAYFA: AUDIT LOG
# ============================================================
elif page == "🕓 Değişiklik Geçmişi":
    st.title("Değişiklik Geçmişi (Audit Log)")
    try:
        df_audit = audit_read(project_name, 500)
        if not df_audit.empty:
            st.dataframe(df_audit, use_container_width=True, hide_index=True)
            st.metric("Toplam Kayıt", len(df_audit))
        else:
            st.info("Bu proje için henüz değişiklik kaydı yok.")
    except Exception as e:
        st.warning(f"Audit log okunamadı: {e}")

# ============================================================
# SAYFA: SİSTEM
# ============================================================
elif page == "⚙️ Sistem ve Yedek":
    st.title("Sistem ve Veri Yönetimi")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 💾 Snapshot Kaydet")
        if st.button("💾 Snapshot kaydet", use_container_width=True):
            payload = {
                "plan": plan.to_dict(orient="records"),
                "actual": actual.to_dict(orient="records"),
                "schedule": schedule.to_dict(orient="records"),
                "risk": risk.to_dict(orient="records"),
                "saved_at": datetime.now().isoformat(timespec="seconds"),
            }
            db_save_snapshot(project_name, payload)
            st.success("✅ Snapshot kaydedildi.")
    with c2:
        st.markdown("### 🔄 Snapshot Geri Yükle")
        if st.button("📂 Snapshot'tan geri yükle", use_container_width=True):
            snap = db_load_snapshot(project_name)
            if snap:
                st.session_state.plan = pd.DataFrame(snap["plan"], columns=PLAN_COLS)
                st.session_state.actual = pd.DataFrame(snap["actual"], columns=ACTUAL_COLS)
                st.session_state.schedule = pd.DataFrame(snap["schedule"], columns=SCHEDULE_COLS)
                st.session_state.risk = pd.DataFrame(snap["risk"], columns=RISK_COLS)
                st.success("✅ Geri yüklendi."); st.rerun()
            else:
                st.warning("Kayıtlı snapshot yok.")

    st.markdown("---")
    st.markdown("### 📊 DB Tablo Durumları")
    tablolar = ["purchase_records","labor_records","progress_records",
                "stock_records","cashflow_records","project_snapshots","audit_log"]
    dur = []
    for t in tablolar:
        try:
            with db_connect() as con:
                c = con.execute(f"SELECT COUNT(*) AS c FROM {t}").fetchone()["c"]
            dur.append({"Tablo": t, "Satır Sayısı": int(c)})
        except Exception:
            dur.append({"Tablo": t, "Satır Sayısı": -1})
    st.dataframe(pd.DataFrame(dur), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### ⚠️ Tehlikeli İşlemler")
    with st.expander("🗑️ Demo verisini sıfırla"):
        if st.button("Session verisini sıfırla", type="secondary"):
            for k in ["plan","actual","schedule","risk"]:
                st.session_state.pop(k, None)
            st.success("Sıfırlandı."); st.rerun()

# ============================================================
# FOOTER
# ============================================================
st.markdown(f"""
<div class="ea-footer">
  <strong>EMİR ASEL İNŞAAT</strong> · Proje Kontrol Sistemi V{APP_VERSION}<br>
  Yerel kayıt: SQLite ({DB_PATH.name}) · Rapor: Excel + PDF + JSON<br>
  <span style="color:{RED_ACCENT};">●</span> Karar destek amaçlıdır; resmi hakediş/muhasebe yerine geçmez.
</div>
""", unsafe_allow_html=True)
