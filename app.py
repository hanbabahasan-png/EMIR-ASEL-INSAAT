# -*- coding: utf-8 -*-
"""EMİR ASEL İNŞAAT | V14.5 — Tek dosya: ERP + YFK + Fatura Öğrenme + Revizyon + Anomali + Monte Carlo + AI Asistan"""
from __future__ import annotations
import os, re, json, uuid, math, hmac, hashlib, mimetypes, tempfile, subprocess, difflib
from io import BytesIO
from pathlib import Path
from urllib.parse import urljoin, urlparse
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from typing import Any
import numpy as np, pandas as pd, streamlit as st
import plotly.express as px, plotly.graph_objects as go
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

APP_NAME="EMİR ASEL İNŞAAT"; APP_VERSION="14.5"; TZ=ZoneInfo("Europe/Istanbul")
ROOT=Path(__file__).resolve().parent; DATA_DIR=ROOT/"data"; UPLOAD_DIR=DATA_DIR/"uploads"
try:
    from dotenv import load_dotenv
    load_dotenv(ROOT/".env", override=False)
except ImportError: pass
for p in (DATA_DIR, UPLOAD_DIR): p.mkdir(parents=True,exist_ok=True)
MAX_MB=int(os.getenv("EA_MAX_UPLOAD_MB","100")); MAX_BYTES=MAX_MB*1024*1024
CURRENCIES=["TRY","EUR","USD","GBP"]
ROLES=["Yönetici","Proje Müdürü","Teknik Ofis","Şantiye Şefi","Finans / Muhasebe","Satın Alma","Depo / Ambar","Görüntüleyici"]
NAVY="#14263D"; RED="#C83E4D"; TEAL="#278C82"
st.set_page_config(page_title=f"{APP_NAME} V{APP_VERSION}",page_icon="🏗️",layout="wide",initial_sidebar_state="expanded")
st.markdown(f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap');
html,body,[class*="css"]{{font-family:Montserrat,sans-serif}}.stApp{{background:#F6F8FA}}
.ea-head{{background:linear-gradient(135deg,#0A192F,{NAVY});padding:20px 26px;border-radius:14px;color:white;margin-bottom:16px}}
.ea-head h1{{color:white!important;margin:0;font-size:24px;font-weight:800}}.ea-head small{{color:#C4D0DF}}
section[data-testid="stSidebar"]{{background:linear-gradient(180deg,#0A192F,{NAVY})}}section[data-testid="stSidebar"] *{{color:white!important}}
div[data-testid="stMetric"]{{background:white;border:1px solid #E5EAF0;border-left:4px solid {RED};padding:11px;border-radius:9px}}
</style>""",unsafe_allow_html=True)

# ŞEMALAR
PLAN=["_id","WBS","Poz No","İş Kalemi","Ana Kategori","Alt Kategori","Birim","Miktar","Birim Fiyat","Planlanan Tutar","Para Birimi","Kur","Kur Kaynağı","Baz Para Birimi Tutarı","Fiyat Yılı","Fiyat Kaynağı","Başlangıç","Bitiş","Sorumlu","Metraj Durumu","Açıklama"]
ACTUAL=["_id","WBS","Poz No","İş Kalemi","Ana Kategori","Birim","Gerçekleşen Miktar","Gerçekleşen Birim Fiyat","Gerçekleşen Tutar","Para Birimi","Kur","Kur Kaynağı","Baz Para Birimi Tutarı","Tarih","Tedarikçi","Belge No","Kayıt Türü","Durum","Onay Tarihi","Onaylayan","Açıklama"]
SCHEDULE=["_id","WBS","Aktivite","Süre (gün)","Öncül WBS","Bağımlılık Tipi","Başlangıç","Bitiş","Fiili Başlangıç","Fiili Bitiş","Tamamlanma (%)","Ağırlık (%)","Sorumlu","Durum","Kanıt / Saha Notu"]
RISK=["_id","Risk ID","Risk Tanımı","Kategori","Olasılık (1-5)","Etki (1-5)","Risk Skoru","Önlem / Aksiyon","Sorumlu","Durum","Tahmini Maliyet Etkisi","Para Birimi","Hedef Tarih"]
CASH=["_id","Tarih","Tür","Açıklama","Firma / Cari","Tutar","Para Birimi","Kur","Vade Tarihi","Durum","WBS","Belge No"]
LIB=["_id","Poz No","İş Kalemi","Birim","Birim Fiyat","Para Birimi","Fiyat Yılı","Kaynak","Kaynak Dosya","Kaynak Sayfa","Yayın / Güncelleme Tarihi","Aktif","Kaynak Türü","Güncelleme Dönemi","Geçerlilik Başlangıcı","Yayın Tarihi","PDF SHA256","Güven Skoru"]
RAYIC=["_id","Rayiç No","Açıklama","Birim","Rayiç Fiyat","Para Birimi","Fiyat Yılı","Kaynak Türü","Kaynak","Kaynak Dosya","Güncelleme Dönemi","Yayın Tarihi","PDF SHA256","Durum"]
YFK_CAND=["_id","Kabul","Kayıt Grubu","Belge Kategorisi","Poz No","İş Kalemi","Birim","Birim Fiyat","Para Birimi","Fiyat Yılı","Kaynak","Kaynak Dosya","Kaynak Sayfa","Aktif","Güncelleme Dönemi","Yayın Tarihi","PDF URL","PDF SHA256","Güven Skoru","İnceleme Durumu"]
YFK_RUNS=["_id","Güncelleme Dönemi","Belge Kategorisi","PDF URL","PDF SHA256","Aday Sayısı","Durum","Kontrol Tarihi","Not"]
PURCHASE=["_id","Sipariş No","WBS","Poz No","Malzeme / Hizmet","Tedarikçi","Sipariş Tarihi","Termin Tarihi","Sipariş Miktarı","Teslim Miktarı","Birim","Birim Fiyat","Para Birimi","Kur","Sipariş Tutarı","Faturalanan Tutar","Fatura No","Fatura Tarihi","Durum","Onay Durumu","Not"]
LABOR=["_id","Tarih","Ekip","İş Tanımı","Çalışan Sayısı","Gün","Kişi Günlük Ücret","Toplam İşçilik","Para Birimi","WBS","Not"]
INVENTORY=["_id","Tarih","Depo","Malzeme Kodu","Malzeme","Hareket Türü","Miktar","Birim","Birim Fiyat","Para Birimi","Tutar","WBS","Belge No","Açıklama"]
HAKEDIS=["_id","Hakediş No","Dönem","Tarih","Taşeron / Firma","WBS","Brüt Tutar","Kesinti","Net Hakediş","KDV Oranı","KDV Tutarı","Ödenecek Tutar","Ödenen","Para Birimi","Durum","Açıklama"]
PROJECT=["_id","Proje Adı","İşveren","Lokasyon","Sözleşme No","Sözleşme Tarihi","Başlangıç Tarihi","Bitiş Tarihi","Para Birimi","KDV Oranı","Risk Rezervi (%)","Proje Yöneticisi","Açıklama"]
CONTRACT=["_id","Sözleşme No","Sözleşme Türü","Karşı Taraf","Kapsam","Sözleşme Tarihi","Bitiş Tarihi","Orijinal Bedel","Para Birimi","Kur","Avans (%)","Teminat (%)","Durum","Onay Tarihi","Açıklama"]
CHANGE=["_id","Değişiklik No","Sözleşme No","Para Birimi","Talep Tarihi","Açıklama","Gelir Etkisi","Maliyet Etkisi","Süre Etkisi (gün)","Durum","Talep Eden","Açıklama"]
EMP_HAK=["_id","Hakediş No","Dönem","Tarih","Brüt İmalat Tutarı","Avans Mahsubu","Teminat Kesintisi","Diğer Kesintiler","Kesinti Toplamı","KDV Oranı","KDV Tutarı","Net Tahakkuk","Hakediş Toplamı","Tahsil Edilen","Kalan Alacak","Para Birimi","Durum","Açıklama"]
DAILY=["_id","Rapor Tarihi","Hava Durumu","Toplam Personel","Günlük İmalat Özeti","Açıklama","Raporlayan"]
APPROVAL=["_id","Talep Tarihi","Veri Seti","Kayıt ID","Talep Başlığı","Tutar","Para Birimi","Talep Eden","Onay Rolü","Durum","Karar Tarihi","Karar Veren","Karar Notu"]
MEMORY=["_id","Tedarikçi","İş Kalemi","Ana Kategori","Poz No","Kullanım","Son Kullanım"]
DS={"plan":PLAN,"actual":ACTUAL,"schedule":SCHEDULE,"risk":RISK,"cash":CASH,"library":LIB,"rayic":RAYIC,"yfk_candidates":YFK_CAND,"yfk_runs":YFK_RUNS,"purchase":PURCHASE,"labor":LABOR,"inventory":INVENTORY,"hakediş":HAKEDIS,"project":PROJECT,"contracts":CONTRACT,"changes":CHANGE,"employer_hakediş":EMP_HAK,"daily_reports":DAILY,"approvals":APPROVAL,"memory":MEMORY}
DATE_FIELDS={"plan":["Başlangıç","Bitiş"],"actual":["Tarih","Onay Tarihi"],"schedule":["Başlangıç","Bitiş"],"risk":["Hedef Tarih"],"cash":["Tarih","Vade Tarihi"],"purchase":["Sipariş Tarihi","Termin Tarihi","Fatura Tarihi"],"labor":["Tarih"],"inventory":["Tarih"],"hakediş":["Tarih"],"project":["Sözleşme Tarihi","Başlangıç Tarihi","Bitiş Tarihi"],"contracts":["Sözleşme Tarihi","Bitiş Tarihi"],"changes":["Talep Tarihi"],"employer_hakediş":["Tarih"],"daily_reports":["Rapor Tarihi"]}
CATEGORIES=["Proje ve Ruhsat","Hafriyat ve Altyapı","Betonarme İmalatlar","Duvar / Sıva / Şap","Yalıtım / Çatı / Cephe","Mimari İnce İşler","Mekanik ve Elektrik","Peyzaj / Çevre","İşçilik / Taşeron","Satın Alma / Lojistik","Diğer"]
GLOBAL_DATASETS={"library","rayic","yfk_candidates","yfk_runs"}

# YARDIMCILAR
def now_tr(): return datetime.now(TZ)
def uid(): return uuid.uuid4().hex
def _miss(x):
    if x is None: return True
    try: return bool(pd.isna(x))
    except: return False
def num(x,d=0.0):
    if _miss(x): return d
    if isinstance(x,(int,float,np.number)):
        try: return float(x)
        except: return d
    s=str(x).strip().replace("₺","").replace("$","").replace("€","").replace("£","").replace(" ","")
    if not s: return d
    if "," in s and "." in s: s=s.replace(".","").replace(",",".") if s.rfind(",")>s.rfind(".") else s.replace(",","")
    elif "," in s: s=s.replace(",",".")
    try: return float(s)
    except: return d
def norm(x): return re.sub(r"\s+"," ",str(x or "").casefold().replace("\u0307","").translate(str.maketrans({"ı":"i","ş":"s","ğ":"g","ü":"u","ö":"o","ç":"c"}))).strip()
def cur_code(x,default=""):
    if _miss(x): return default
    s=str(x).strip().upper(); a={"₺":"TRY","TL":"TRY","€":"EUR","EURO":"EUR","$":"USD","£":"GBP"}
    return a.get(s,s if s in CURRENCIES else default)
def money(x,c="TRY"):
    if _miss(x): return "—"
    return f"{ {'TRY':'₺','EUR':'€','USD':'$','GBP':'£'}.get(c,c)} {float(x):,.2f}".replace(",","X").replace(".",",").replace("X",".")
def jsafe(x):
    if x is None: return None
    if isinstance(x,(datetime,date,pd.Timestamp)): return None if pd.isna(x) else x.isoformat()
    if isinstance(x,dict): return {str(k):jsafe(v) for k,v in x.items()}
    if isinstance(x,(list,tuple)): return [jsafe(v) for v in x]
    if isinstance(x,np.integer): return int(x)
    if isinstance(x,(float,np.floating)):
        y=float(x); return None if not math.isfinite(y) else y
    if isinstance(x,(np.bool_,bool)): return bool(x)
    if _miss(x): return None
    if isinstance(x,(str,int)): return x
    return str(x)
def pl(d): return json.dumps(jsafe(d),ensure_ascii=False,allow_nan=False)
def ef(df,cols):
    x=pd.DataFrame() if df is None else df.copy()
    for c in cols:
        if c not in x: x[c]=None
    x=x[cols].copy()
    if "_id" in cols:
        x["_id"]=x["_id"].apply(lambda v:str(v).strip() if not _miss(v) and str(v).strip() not in("","nan","None") else uid())
        dup=x["_id"].duplicated()
        if dup.any(): x.loc[dup,"_id"]=[uid() for _ in range(int(dup.sum()))]
    return x.reset_index(drop=True)
def to_dates(s): return pd.to_datetime(s,errors="coerce",dayfirst=True)
def pct_fmt(v): return "—" if _miss(v) else f"%{float(v):.1f}".replace(".",",")
def is_approved(s): return norm(s) in {"onaylandi","approved","yururlukte"}
def approved_mask(s): return s.fillna("").map(is_approved).astype(bool)
def categorize(desc):
    t=norm(desc)
    R=[("Proje ve Ruhsat",["mimari proje","statik proje","ruhsat"]),("Hafriyat ve Altyapı",["hafriyat","kazı","dolgu","iksa","drenaj"]),("Betonarme İmalatlar",["beton","grobeton","radye","kolon","kiriş","perde","kalıp","donatı","demir"]),("Duvar / Sıva / Şap",["duvar","tuğla","sıva","şap"]),("Yalıtım / Çatı / Cephe",["yalıtım","izolasyon","membran","çatı","cephe"]),("Mimari İnce İşler",["seramik","kapı","pencere","boya","parke","mermer"]),("Mekanik ve Elektrik",["mekanik","elektrik","kablo","pano","hvac","yangın"]),("Peyzaj / Çevre",["peyzaj","çevre","bahçe"]),("İşçilik / Taşeron",["işçilik","taşeron","puantaj"]),("Satın Alma / Lojistik",["nakliye","lojistik","tedarik"])]
    for cat,keys in R:
        if any(k in t for k in keys): return cat
    return "Diğer"
def find_col(df,cands):
    nc=lambda c: re.sub(r"[^a-z0-9]+","",norm(c))
    lookup={nc(c):str(c) for c in df.columns}
    for cand in cands:
        if nc(cand) in lookup: return lookup[nc(cand)]
    best=None; bs=0
    for col in df.columns:
        s=max([difflib.SequenceMatcher(None,nc(c),nc(col)).ratio() for c in cands] or [0])
        if s>bs: bs=s; best=str(col)
    return best if bs>=0.85 else None
def fx(amount,src,base,rate):
    a=num(amount,np.nan); s=cur_code(src); b=cur_code(base)
    if not math.isfinite(a) or not s or not b: return np.nan
    if s==b: return a
    r=num(rate,np.nan); return a*r if math.isfinite(r) and r>0 else np.nan
def calc_plan(df,base):
    x=ef(df,PLAN)
    x["Miktar"]=x["Miktar"].map(num); x["Birim Fiyat"]=x["Birim Fiyat"].map(num)
    x["Planlanan Tutar"]=x["Miktar"]*x["Birim Fiyat"]
    x["Baz Para Birimi Tutarı"]=[fx(a,c,base,r) for a,c,r in zip(x["Planlanan Tutar"],x["Para Birimi"],x["Kur"])]
    return x
def calc_actual(df,base):
    x=ef(df,ACTUAL)
    x["Gerçekleşen Miktar"]=x["Gerçekleşen Miktar"].map(num); x["Gerçekleşen Birim Fiyat"]=x["Gerçekleşen Birim Fiyat"].map(num)
    x["Gerçekleşen Tutar"]=x["Gerçekleşen Miktar"]*x["Gerçekleşen Birim Fiyat"]
    x["Baz Para Birimi Tutarı"]=[fx(a,c,base,r) for a,c,r in zip(x["Gerçekleşen Tutar"],x["Para Birimi"],x["Kur"])]
    return x

# VERİTABANI
def db_url():
    u=os.getenv("DATABASE_URL","").strip()
    if not u:
        try: u=str(st.secrets.get("DATABASE_URL","")).strip()
        except: u=""
    if not u: return f"sqlite:///{(ROOT/'emir_asel_v12.db').as_posix()}"
    if u.startswith("postgres://"): return "postgresql+psycopg://"+u[11:]
    if u.startswith("postgresql://"): return "postgresql+psycopg://"+u[13:]
    return u
@st.cache_resource(show_spinner=False)
def get_engine():
    u=db_url(); args={"check_same_thread":False,"timeout":30} if u.startswith("sqlite") else {}
    return create_engine(u,pool_pre_ping=True,future=True,connect_args=args)
try: ENGINE=get_engine(); DB_KIND=ENGINE.dialect.name
except Exception as e: ENGINE=None; DB_KIND="unknown"; DB_ERR=str(e)
else: DB_ERR=""
def init_db():
    if ENGINE is None: raise RuntimeError(DB_ERR)
    with ENGINE.begin() as con:
        con.execute(text("CREATE TABLE IF NOT EXISTS ea_records(project_key TEXT NOT NULL,dataset TEXT NOT NULL,record_id TEXT NOT NULL,payload TEXT NOT NULL,updated_at TEXT NOT NULL,PRIMARY KEY(project_key,dataset,record_id))"))
        con.execute(text("CREATE INDEX IF NOT EXISTS ix_rec ON ea_records(project_key,dataset)"))
        con.execute(text("CREATE TABLE IF NOT EXISTS ea_users(username TEXT PRIMARY KEY,display_name TEXT NOT NULL,role TEXT NOT NULL,salt TEXT NOT NULL,password_hash TEXT NOT NULL,active INTEGER NOT NULL DEFAULT 1,created_at TEXT NOT NULL)"))
        con.execute(text("CREATE TABLE IF NOT EXISTS ea_file_assets(id TEXT PRIMARY KEY,project_key TEXT NOT NULL,original_name TEXT NOT NULL,stored_name TEXT NOT NULL,extension TEXT NOT NULL,sha256 TEXT NOT NULL,size_bytes INTEGER NOT NULL,uploaded_by TEXT NOT NULL,created_at TEXT NOT NULL,UNIQUE(project_key,sha256))"))
def load_ds(proj,ds,cols):
    with ENGINE.connect() as con: rows=con.execute(text("SELECT payload FROM ea_records WHERE project_key=:p AND dataset=:d ORDER BY updated_at,record_id"),{"p":proj,"d":ds}).fetchall()
    recs=[]
    for r in rows:
        try: recs.append(json.loads(r[0]))
        except: pass
    out=ef(pd.DataFrame(recs),cols)
    for c in DATE_FIELDS.get(ds,[]):
        if c in out: out[c]=pd.to_datetime(out[c],errors="coerce",dayfirst=True).dt.date
    return out
def save_ds(proj,ds,df,cols):
    frame=ef(df,cols); stamp=now_tr().isoformat(timespec="seconds")
    stmt=text("INSERT INTO ea_records(project_key,dataset,record_id,payload,updated_at) VALUES(:p,:d,:r,:v,:t) ON CONFLICT(project_key,dataset,record_id) DO UPDATE SET payload=excluded.payload,updated_at=excluded.updated_at")
    with ENGINE.begin() as con:
        cur=con.execute(text("SELECT record_id FROM ea_records WHERE project_key=:p AND dataset=:d"),{"p":proj,"d":ds}).fetchall()
        cur_ids={str(r[0]) for r in cur}; keep=set(frame["_id"].astype(str)) if "_id" in frame else set()
        for s in cur_ids-keep: con.execute(text("DELETE FROM ea_records WHERE project_key=:p AND dataset=:d AND record_id=:r"),{"p":proj,"d":ds,"r":s})
        for _,row in frame.iterrows(): con.execute(stmt,{"p":proj,"d":ds,"r":str(row["_id"]),"v":pl(row.to_dict()),"t":stamp})

# AUTH
def pwd_hash(p,s): return hashlib.pbkdf2_hmac("sha256",p.encode(),bytes.fromhex(s),310000).hex()
def list_users():
    with ENGINE.connect() as con: rows=con.execute(text("SELECT username,display_name,role,active,created_at FROM ea_users ORDER BY username")).fetchall()
    return pd.DataFrame([dict(r._mapping) for r in rows]) if rows else pd.DataFrame(columns=["username","display_name","role","active","created_at"])
def create_user(un,dn,role,pw):
    un=un.strip().lower()
    if not re.fullmatch(r"[a-z0-9_.-]{3,40}",un): raise ValueError("Kullanıcı adı 3-40 karakter.")
    if role not in ROLES or len(pw)<12: raise ValueError("Rol veya parola (12+) hatalı.")
    s=os.urandom(16).hex()
    with ENGINE.begin() as con: con.execute(text("INSERT INTO ea_users(username,display_name,role,salt,password_hash,active,created_at) VALUES(:u,:d,:r,:s,:h,1,:t)"),{"u":un,"d":dn.strip() or un,"r":role,"s":s,"h":pwd_hash(pw,s),"t":now_tr().isoformat(timespec="seconds")})
def auth_user(un,pw):
    un=un.strip().lower()
    with ENGINE.connect() as con: r=con.execute(text("SELECT username,display_name,role,salt,password_hash,active FROM ea_users WHERE username=:u"),{"u":un}).first()
    if not r or int(r.active)!=1: return None
    if not hmac.compare_digest(pwd_hash(pw,r.salt),r.password_hash): return None
    return {"username":r.username,"display_name":r.display_name,"role":r.role}
def auth_gate():
    uf=list_users()
    if uf.empty:
        token=os.getenv("EA_BOOTSTRAP_TOKEN","").strip()
        st.title("🔐 İlk Kurulum — Yönetici")
        if token: st.info("Güvenli kurulum aktif: kurulum anahtarı gerekli.")
        else: st.warning("⚠️ EA_BOOTSTRAP_TOKEN yok — yerel kurulum için normal. Üretimde tanımlayın.")
        with st.form("bootstrap"):
            supplied=st.text_input("Kurulum anahtarı",type="password") if token else "__NO_TOKEN__"
            un=st.text_input("Kullanıcı adı",value="admin"); dn=st.text_input("Ad Soyad",value="Yönetici")
            p1=st.text_input("Parola (12+)",type="password"); p2=st.text_input("Tekrar",type="password")
            sub=st.form_submit_button("Yönetici Oluştur",type="primary")
        if sub:
            if token and not hmac.compare_digest(supplied,token): st.error("Anahtar hatalı.")
            elif p1!=p2: st.error("Parolalar eşleşmiyor.")
            elif len(p1)<12: st.error("Parola 12+ olmalı.")
            else:
                try: create_user(un,dn,"Yönetici",p1); st.success("Yönetici oluşturuldu."); st.rerun()
                except Exception as e: st.error(str(e))
        st.stop()
    if not st.session_state.get("auth_user"):
        st.title("🔐 EMİR ASEL — Giriş")
        with st.form("login"):
            un=st.text_input("Kullanıcı adı"); pw=st.text_input("Parola",type="password")
            sub=st.form_submit_button("Giriş",type="primary")
        if sub:
            u=auth_user(un,pw)
            if u: st.session_state.auth_user=u; st.rerun()
            else: st.error("Hatalı giriş.")
        st.stop()
    cur=st.session_state.auth_user; fresh=uf[uf.username.astype(str)==str(cur.get("username"))]
    if fresh.empty or int(fresh.iloc[0].active)!=1: st.session_state.pop("auth_user",None); st.error("Hesap pasif."); st.stop()
    cur["role"]=str(fresh.iloc[0].role)

# CPM
def cpm_calc(sched):
    if sched.empty: return pd.DataFrame()
    d=ef(sched,SCHEDULE); d["WBS"]=d.WBS.fillna("").astype(str).str.strip(); d["Süre (gün)"]=pd.to_numeric(d["Süre (gün)"],errors="coerce").fillna(0).clip(lower=0)
    d=d[d.WBS.ne("")].copy()
    if d.WBS.duplicated().any(): return d
    dur=dict(zip(d.WBS,d["Süre (gün)"])); succ={k:[] for k in dur}; indeg={k:0 for k in dur}; pred={}
    for _,r in d.iterrows():
        edges=[]
        for tok in re.split(r"[,;]",str(r.get("Öncül WBS",""))):
            p=tok.strip().split(":")
            if not p[0]: continue
            edges.append((p[0],(p[1] if len(p)>1 else "FS").upper(),num(p[2]) if len(p)>2 else 0))
        pred[r.WBS]=edges
        for par,_,_ in edges:
            if par in dur: succ[par].append((r.WBS,None,None)); indeg[r.WBS]+=1
    q=[k for k,v in indeg.items() if v==0]; order=[]
    while q:
        k=q.pop(0); order.append(k)
        for nxt,_,_ in succ[k]:
            indeg[nxt]-=1
            if indeg[nxt]==0: q.append(nxt)
    if len(order)!=len(dur): return d
    es={}; ef={}
    for k in order:
        n=dur[k]; lim=[1.0]
        for par,rel,lag in pred[k]:
            if par not in ef: continue
            if rel=="FS": lim.append(ef[par]+lag+1)
            elif rel=="SS": lim.append(es[par]+lag)
            elif rel=="FF": lim.append(ef[par]+lag-n+1)
            else: lim.append(es[par]+lag-n+1)
        es[k]=max(lim); ef[k]=es[k]+max(n,0)-1
    fin=max(ef.values(),default=1); ls={}; lf={}
    for k in reversed(order):
        n=dur[k]; lim=[]
        for nxt,rel,lag in succ[k]:
            if nxt not in ls: continue
            if rel=="FS": lim.append(ls[nxt]-lag-1)
            elif rel=="SS": lim.append(ls[nxt]-lag+n-1)
            elif rel=="FF": lim.append(lf[nxt]-lag)
            else: lim.append(lf[nxt]-lag+n-1)
        lf[k]=min(lim) if lim else fin; ls[k]=lf[k]-max(n,0)+1
    d["Erken Başlangıç"]=d.WBS.map(es); d["Erken Bitiş"]=d.WBS.map(ef); d["Bolluk (gün)"]=d.WBS.map(ls)-d.WBS.map(es); d["Kritik mi?"]=np.where(d["Bolluk (gün)"].abs()<1e-8,"EVET","Hayır")
    return d

# EVM
def evm_calc(plan,actual,sched,report_date,base):
    p=calc_plan(plan,base); a=calc_actual(actual,base)
    pp=p[approved_mask(p["Metraj Durumu"])]; aa=a[approved_mask(a["Durum"])]
    bac=float(pd.to_numeric(pp["Baz Para Birimi Tutarı"],errors="coerce").fillna(0).sum()) if not pp.empty else 0
    ad=to_dates(aa["Tarih"]) if not aa.empty else pd.Series(dtype="datetime64[ns]")
    ontime=ad.notna()&(ad.dt.date<=report_date) if not aa.empty else pd.Series(dtype=bool)
    af=aa.loc[ontime] if not aa.empty else aa
    ac=float(pd.to_numeric(af["Baz Para Birimi Tutarı"],errors="coerce").fillna(0).sum()) if not af.empty else 0
    pv=0; ev=0
    if not pp.empty and not sched.empty:
        s=ef(sched,SCHEDULE); s["Tamamlanma (%)"]=pd.to_numeric(s["Tamamlanma (%)"],errors="coerce").fillna(0).clip(0,100)
        s["Ağırlık (%)"]=pd.to_numeric(s["Ağırlık (%)"],errors="coerce").fillna(0)
        s["Süre (gün)"]=pd.to_numeric(s["Süre (gün)"],errors="coerce").fillna(0)
        s["Başlangıç"]=to_dates(s["Başlangıç"]); s["Bitiş"]=to_dates(s["Bitiş"])
        prog={}; planprog={}
        for w,grp in s.groupby("WBS"):
            wts=grp["Ağırlık (%)"].copy()
            if wts.sum()<=0: wts=grp["Süre (gün)"]
            prog[w]=float(np.average(grp["Tamamlanma (%)"],weights=wts))/100 if wts.sum()>0 else float(grp["Tamamlanma (%)"].mean())/100
            fracs=[]; wts2=[]
            for _,t in grp.iterrows():
                st_dt,en_dt=t["Başlangıç"],t["Bitiş"]
                if pd.isna(st_dt) or pd.isna(en_dt) or en_dt<st_dt: continue
                rd=pd.Timestamp(report_date)
                f=0 if rd<st_dt else (1 if rd>=en_dt else min(max((rd-st_dt).days/max((en_dt-st_dt).days,1),0),1))
                fracs.append(f); tw=num(t["Ağırlık (%)"],0) or num(t["Süre (gün)"],0) or 1; wts2.append(tw)
            if fracs: planprog[w]=float(np.average(fracs,weights=wts2))
        for _,r in pp.iterrows():
            amt=num(r.get("Baz Para Birimi Tutarı"),np.nan)
            if not math.isfinite(amt): continue
            w=str(r.get("WBS","")).strip(); ev+=amt*prog.get(w,0); pv+=amt*planprog.get(w,0)
    cpi=ev/ac if ac>0 else np.nan; spi=ev/pv if pv>0 else np.nan; eac=bac/cpi if pd.notna(cpi) and cpi>0 else np.nan
    return {"BAC":bac,"PV":pv,"EV":ev,"AC":ac,"CV":ev-ac,"SV":ev-pv,"CPI":cpi,"SPI":spi,"EAC":eac,"VAC":bac-eac if pd.notna(eac) else np.nan}

def wprogress(s):
    if s.empty: return 0.0
    p=pd.to_numeric(s["Tamamlanma (%)"],errors="coerce").fillna(0).clip(0,100); w=pd.to_numeric(s["Ağırlık (%)"],errors="coerce").fillna(0).clip(lower=0)
    if w.sum()<=0: w=pd.to_numeric(s["Süre (gün)"],errors="coerce").fillna(0).clip(lower=0)
    return float(np.average(p,weights=w)) if w.sum()>0 else float(p.mean())

def emp_hak_calc(df):
    x=ef(df,EMP_HAK)
    for c in ["Brüt İmalat Tutarı","Avans Mahsubu","Teminat Kesintisi","Diğer Kesintiler","Tahsil Edilen"]: x[c]=x[c].map(num)
    vat=x["KDV Oranı"].map(num); x["KDV Oranı"]=np.where(vat>1,vat/100,vat).clip(0,1)
    x["Kesinti Toplamı"]=x["Avans Mahsubu"]+x["Teminat Kesintisi"]+x["Diğer Kesintiler"]
    x["Net Tahakkuk"]=x["Brüt İmalat Tutarı"]-x["Kesinti Toplamı"]
    x["KDV Tutarı"]=x["Net Tahakkuk"]*x["KDV Oranı"]; x["Hakediş Toplamı"]=x["Net Tahakkuk"]+x["KDV Tutarı"]
    dlt=x["Hakediş Toplamı"]-x["Tahsil Edilen"]; x["Kalan Alacak"]=dlt.clip(lower=0)
    return x

def summary_costs(plan,actual,base):
    p=calc_plan(plan,base); a=calc_actual(actual,base)
    p=p[approved_mask(p["Metraj Durumu"])]; a=a[approved_mask(a["Durum"])]
    cats=sorted(set(p["Ana Kategori"].dropna().astype(str))|set(a["Ana Kategori"].dropna().astype(str)))
    rows=[]
    for cat in cats:
        pb=float(pd.to_numeric(p[p["Ana Kategori"].fillna("")==cat]["Baz Para Birimi Tutarı"],errors="coerce").fillna(0).sum())
        ab=float(pd.to_numeric(a[a["Ana Kategori"].fillna("")==cat]["Baz Para Birimi Tutarı"],errors="coerce").fillna(0).sum())
        rows.append({"Ana Kategori":cat,"Bütçe":pb,"Gerçekleşen":ab,"Sapma":ab-pb,"Sapma (%)":((ab-pb)/pb*100 if pb else 0)})
    return pd.DataFrame(rows).sort_values("Bütçe",ascending=False)

# ============ V14.5 ÖZELLİKLERİ ============

# 1. FATURA ÖĞRENME
def learn_mapping(proj,ted,item,cat,poz,mem_df):
    df=mem_df.copy() if not mem_df.empty else pd.DataFrame(columns=MEMORY)
    m=(df["Tedarikçi"].astype(str).str.strip()==str(ted).strip())&(df["İş Kalemi"].astype(str).str.strip()==str(item).strip())
    if m.any():
        df.loc[m,"Kullanım"]=pd.to_numeric(df.loc[m,"Kullanım"],errors="coerce").fillna(0)+1
        df.loc[m,"Son Kullanım"]=now_tr().isoformat()
        df.loc[m,"Ana Kategori"]=cat; df.loc[m,"Poz No"]=poz
    else:
        df=pd.concat([df,pd.DataFrame([{"_id":uid(),"Tedarikçi":ted,"İş Kalemi":item,"Ana Kategori":cat,"Poz No":poz,"Kullanım":1,"Son Kullanım":now_tr().isoformat()}])],ignore_index=True)
    return df

def suggest_from_memory(ted,item,mem_df):
    if mem_df.empty: return None
    t=norm(ted); k=norm(item); best=None; bs=0
    for _,r in mem_df.iterrows():
        sc=0
        if norm(r.get("Tedarikçi",""))==t: sc+=0.6
        km=norm(r.get("İş Kalemi",""))
        if km and k:
            common=len(set(k.split())&set(km.split())); total=max(len(set(k.split())|set(km.split())),1)
            sc+=common/total*0.4
        if sc>bs: bs=sc; best=r
    return best.to_dict() if bs>=0.5 else None

# 2. REVİZYON KARŞILAŞTIRMA
def compare_revisions(a,b,key_cols=None,val_cols=None):
    key_cols=key_cols or ["Poz No","İş Kalemi"]; val_cols=val_cols or ["Miktar","Birim Fiyat"]
    A=a.copy() if a is not None else pd.DataFrame(); B=b.copy() if b is not None else pd.DataFrame()
    for c in key_cols:
        if c not in A: A[c]=""
        if c not in B: B[c]=""
    for c in val_cols:
        if c not in A: A[c]=0
        if c not in B: B[c]=0
    A["_k"]=A[key_cols].astype(str).agg("|".join,axis=1); B["_k"]=B[key_cols].astype(str).agg("|".join,axis=1)
    m=A.merge(B,on="_k",how="outer",suffixes=("_E","_Y"),indicator=True)
    out=pd.DataFrame()
    out["Poz No"]=m.get("Poz No_Y",m.get("Poz No_E",""))
    out["İş Kalemi"]=m.get("İş Kalemi_Y",m.get("İş Kalemi_E",""))
    for c in val_cols:
        out[f"{c} (E)"]=pd.to_numeric(m.get(f"{c}_E",0),errors="coerce").fillna(0)
        out[f"{c} (Y)"]=pd.to_numeric(m.get(f"{c}_Y",0),errors="coerce").fillna(0)
        out[f"{c} Fark"]=out[f"{c} (Y)"]-out[f"{c} (E)"]
    if "Miktar Fark" in out and "Birim Fiyat Fark" in out:
        out["Tutar Farkı"]=out["Miktar (Y)"]*out["Birim Fiyat (Y)"]-out["Miktar (E)"]*out["Birim Fiyat (E)"]
    out["Durum"]=m["_merge"].map({"left_only":"SİLİNDİ","right_only":"EKLENDİ","both":"DEĞİŞTİ"})
    same=(out["Durum"]=="DEĞİŞTİ")&(out.get("Miktar Fark",0).abs()<1e-6)&(out.get("Birim Fiyat Fark",0).abs()<1e-6)
    out.loc[same,"Durum"]="AYNI"
    return out[out["Durum"]!="AYNI"].reset_index(drop=True)

# 3. ANOMALİ
def detect_anomalies(df,iqr=1.5):
    out=[]
    if df is None or df.empty: return pd.DataFrame(columns=["Tür","Alan","Poz No","İş Kalemi","Değer","Yön","Önem"])
    for col in ["Miktar","Birim Fiyat","Gerçekleşen Miktar","Gerçekleşen Birim Fiyat"]:
        if col not in df.columns: continue
        vals=pd.to_numeric(df[col],errors="coerce").fillna(0); nz=vals[vals>0]
        if len(nz)<5: continue
        q1,q3=nz.quantile([.25,.75]); rg=q3-q1
        if rg<=0: continue
        lo,hi=q1-iqr*rg,q3+iqr*rg
        for idx in vals.index[(vals>0)&((vals<lo)|(vals>hi))]:
            out.append({"Tür":"Aykırı Değer","Alan":col,"Poz No":str(df.at[idx,"Poz No"]) if "Poz No" in df else "","İş Kalemi":str(df.at[idx,"İş Kalemi"])[:60] if "İş Kalemi" in df else "","Değer":float(vals[idx]),"Yön":"Yüksek" if vals[idx]>hi else "Düşük","Önem":"Yüksek" if abs(vals[idx]-hi)/max(hi,1)>0.5 else "Orta"})
    if "Başlangıç" in df.columns and "Bitiş" in df.columns:
        sd=to_dates(df["Başlangıç"]); ed=to_dates(df["Bitiş"])
        bad=(ed<sd)&ed.notna()&sd.notna()
        for idx in df.index[bad]:
            out.append({"Tür":"Tarih Mantıksız","Alan":"Tarih","Poz No":str(df.at[idx,"Poz No"]) if "Poz No" in df else "","İş Kalemi":str(df.at[idx,"İş Kalemi"])[:60] if "İş Kalemi" in df else "","Değer":"Bitiş<Başlangıç","Yön":"-","Önem":"Yüksek"})
    return pd.DataFrame(out) if out else pd.DataFrame(columns=["Tür","Alan","Poz No","İş Kalemi","Değer","Yön","Önem"])

# 4. MONTE CARLO
def monte_carlo_eac(bac,ac,ev,risk_df=None,it=1000,seed=42):
    if bac<=0: return None
    rng=np.random.default_rng(seed)
    var=0.15
    if risk_df is not None and not risk_df.empty:
        try:
            rs=pd.to_numeric(risk_df.get("Risk Skoru",0),errors="coerce").fillna(0)
            if len(rs): var=min(rs.max()/25*0.30+0.10,0.40)
        except: pass
    sims=[]
    for _ in range(it):
        sc=bac*(1+rng.normal(0,var))
        if ac>0 and ev>0:
            cpi=ev/ac; ec=bac/max(cpi,0.5); eac=sc*0.7+ec*0.3
        else: eac=sc
        sims.append(max(eac,ac))
    sims=np.array(sims)
    return {"P10":float(np.percentile(sims,10)),"P50":float(np.percentile(sims,50)),"P90":float(np.percentile(sims,90)),"Ort":float(sims.mean()),"Sims":sims}

# 5. AI
def ai_client():
    k=os.getenv("OPENAI_API_KEY","").strip()
    if not k:
        try: k=str(st.secrets.get("OPENAI_API_KEY","")).strip()
        except: k=""
    if not k: raise RuntimeError("OPENAI_API_KEY yok.")
    try: from openai import OpenAI
    except: raise RuntimeError("openai paketi yok.")
    return OpenAI(api_key=k)
def ai_image(bts,mime,ctx=""):
    import base64
    c=ai_client(); b64=base64.b64encode(bts).decode(); m=os.getenv("OPENAI_MODEL") or "gpt-4o-mini"
    r=c.chat.completions.create(model=m,messages=[{"role":"user","content":[{"type":"text","text":ctx+"\nGörseli analiz et."},{"type":"image_url","image_url":{"url":f"data:{mime};base64,{b64}"}}]}],max_tokens=1500)
    return r.choices[0].message.content
def ai_text(text,ctx=""):
    c=ai_client(); m=os.getenv("OPENAI_MODEL") or "gpt-4o-mini"
    r=c.chat.completions.create(model=m,messages=[{"role":"system","content":"İnşaat teknik ofis asistanı. Özetle, uydurma."},{"role":"user","content":ctx+"\n\n"+text[:50000]}],max_tokens=1500)
    return r.choices[0].message.content

# YFK
YFK_INDEX="https://yfk.csb.gov.tr/aylik-guncel-rayic-ve-birim-fiyat-listeleri-113351"
YFK_FALLBACK="https://yfk.csb.gov.tr/2026-yilina-ait-guncellenmis-aylik-rayic-ve-birim-fiyat-listeleri-kasim-ayi-03-kasim-2026-tarihinde-yayinlanmis-olup-01-kasim-2026-tarihinden-itibaren-gecerlidir-115099"
YFK_HOSTS={"yfk.csb.gov.tr","webdosya.csb.gov.tr"}
YFK_MONTHS={"ocak":1,"subat":2,"mart":3,"nisan":4,"mayis":5,"haziran":6,"temmuz":7,"agustos":8,"eylul":9,"ekim":10,"kasim":11,"aralik":12}
YFK_TYPES=[("İnşaat Rayiçleri","Rayiç","insaatrayic"),("Mekanik Rayiçleri","Rayiç","mekanikrayic"),("Elektrik Rayiçleri","Rayiç","elektrikrayic"),("İnşaat Birim Fiyatları","Birim Fiyat","insaatbirimfiyat"),("Mekanik Birim Fiyatları","Birim Fiyat","mekanikbirimfiyat"),("Elektrik Birim Fiyatları","Birim Fiyat","elektrikbirimfiyat")]
YFK_UNITS={"m","mt","m2","m3","kg","ton","t","adet","ad","saat","gun","set","takim","lt","l","km","%","kwsaat","kw","kwh"}
YFK_CODE_RE=re.compile(r"^(?:[A-ZÇĞİÖŞÜ]{1,5}[./-]?)?\d{1,4}(?:[./-]\d{1,6}){1,6}$",re.I)

def yfk_req(url,pdf=False):
    try: import requests
    except: raise RuntimeError("requests yok.")
    p=urlparse(url)
    if p.scheme!="https" or p.hostname not in YFK_HOSTS: raise ValueError(f"Güvenlik: {p.hostname}")
    h={"User-Agent":"EmirAsel-YFK/1.0"}
    if pdf:
        mx=int(os.getenv("EA_YFK_MAX_PDF_MB","70"))*1024*1024
        with requests.get(url,headers=h,timeout=(10,35),allow_redirects=True,stream=True) as r:
            r.raise_for_status(); f=urlparse(r.url)
            if f.scheme!="https" or f.hostname not in YFK_HOSTS: raise ValueError("Yönlendirme engellendi.")
            chunks=[]; tot=0
            for ch in r.iter_content(chunk_size=1024*1024):
                if not ch: continue
                tot+=len(ch)
                if tot>mx: raise ValueError("PDF büyük.")
                chunks.append(ch)
            b=b"".join(chunks)
            if not b.startswith(b"%PDF-"): raise ValueError("PDF değil.")
            return type("R",(),{"content":b,"url":r.url})()
    r=requests.get(url,headers=h,timeout=(10,25),allow_redirects=True); r.raise_for_status(); f=urlparse(r.url)
    if f.scheme!="https" or f.hostname not in YFK_HOSTS: raise ValueError("Yönlendirme.")
    return r

def yfk_discover():
    today=now_tr().date()
    try:
        from bs4 import BeautifulSoup
        soup=BeautifulSoup(yfk_req(YFK_INDEX).text,"html.parser"); found=[]
        for a in soup.find_all("a",href=True):
            lbl=a.get_text(" ",strip=True); href=urljoin(YFK_INDEX,a["href"])
            month=None
            for n,v in YFK_MONTHS.items():
                if n in norm(lbl+" "+href): month=v; break
            yr_m=re.search(r"20\d{2}",lbl+" "+href); yr=int(yr_m.group(0)) if yr_m else today.year
            if not month or month>today.month or yr>today.year: continue
            if urlparse(href).hostname!="yfk.csb.gov.tr": continue
            if href.rstrip("/")==YFK_INDEX.rstrip("/"): continue
            found.append((yr,month,href,lbl))
        if found:
            found.sort(key=lambda x:(x[0],x[1]),reverse=True); yr,mo,url,_=found[0]
            return yfk_parse_month(url,yr,mo)
    except Exception as e:
        fb=os.getenv("EA_YFK_FALLBACK_PAGE",YFK_FALLBACK if today.year==2026 and today.month>=11 else "").strip()
        if fb:
            mo=None
            for n,v in YFK_MONTHS.items():
                if n in norm(fb): mo=v; break
            mo=mo or today.month; yr_m=re.search(r"20\d{2}",fb); yr=int(yr_m.group(0)) if yr_m else today.year
            return yfk_parse_month(fb,yr,mo,str(e))
        raise RuntimeError(f"YFK: {e}")
    raise RuntimeError("YFK dizin yok.")

def yfk_parse_month(url,yr,mo,index_err=""):
    try: from bs4 import BeautifulSoup
    except: raise RuntimeError("bs4 yok.")
    r=yfk_req(url); soup=BeautifulSoup(r.text,"html.parser")
    h=soup.find("h1"); title=h.get_text(" ",strip=True) if h else (soup.title.get_text(" ",strip=True) if soup.title else url)
    pub=None; vf=None
    dm=re.search(r"(\d{1,2})\s+([A-Za-zÇĞİÖŞÜçğıöşü]+)\s+(20\d{2}).{0,80}?yayınlanmış olup\s+(\d{1,2})\s+([A-Za-zÇĞİÖŞÜçğıöşü]+)\s+(20\d{2})",title,re.I)
    if dm:
        try:
            pub=date(int(dm.group(3)),YFK_MONTHS[norm(dm.group(2))],int(dm.group(1)))
            vf=date(int(dm.group(6)),YFK_MONTHS[norm(dm.group(5))],int(dm.group(4)))
        except: pass
    docs={}
    for a in soup.find_all("a",href=True):
        lbl=a.get_text(" ",strip=True); href=urljoin(url,a["href"]); nc=re.sub(r"[^a-z0-9]+","",norm(lbl+" "+href))
        if urlparse(href).hostname not in YFK_HOSTS: continue
        for disp,grp,key in YFK_TYPES:
            if key in nc and href.lower().endswith(".pdf"): docs[key]={"cat":disp,"group":grp,"url":href}; break
    period=f"{yr}-{mo:02d}"
    return {"url":url,"title":title,"period":period,"pub":pub,"vf":vf,"docs":docs,"missing":[x[0] for x in YFK_TYPES if x[2] not in docs],"index_err":str(index_err)}

def yfk_num(v):
    s=str(v or "").strip().replace("₺","").replace("TL","").replace(" ","")
    if not s: return np.nan
    if "," in s and "." in s: s=s.replace(".","").replace(",",".") if s.rfind(",")>s.rfind(".") else s.replace(",","")
    elif "," in s: s=s.replace(",",".")
    try:
        x=float(s); return x if math.isfinite(x) else np.nan
    except: return np.nan

def yfk_row(cells,pn,cat,grp,period,pub,vf,src,url,dig):
    cells=[re.sub(r"\s+"," ",str(c or "")).strip() for c in cells]; cells=[c for c in cells if c]
    if len(cells)<2: return None
    n=norm(" ".join(cells))
    if any(k in n for k in ["sayfa no","poz numarasi","birim fiyat listesi","genel toplam","ara toplam"]): return None
    code=""; ci=None
    for i,c in enumerate(cells):
        tk=c.strip().replace(" ","")
        if YFK_CODE_RE.match(tk): code=tk; ci=i; break
    unit=""; ui=None
    for i,c in enumerate(cells):
        u=norm(c).replace("²","2").replace("³","3").replace("ğ","g")
        if u in YFK_UNITS or u.replace(".","") in YFK_UNITS: unit=c.strip(); ui=i; break
    nums=[]
    for i,c in enumerate(cells):
        if i==ci: continue
        v=yfk_num(c)
        if math.isfinite(v) and v>0 and re.fullmatch(r"[\d.,]+",c.replace(" ","")): nums.append((i,v))
    if not nums: return None
    pi,price=nums[-1]; ex={i for i in [ci,ui,pi] if i is not None}
    desc=" ".join(c for i,c in enumerate(cells) if i not in ex and not re.fullmatch(r"[\d.,]+",c.replace(" ",""))).strip(" -|;:")
    if len(desc)<5: return None
    conf=0.97 if code and unit else (0.82 if unit else 0.58)
    return {"_id":uid(),"Kabul":False,"Kayıt Grubu":grp,"Belge Kategorisi":cat,"Poz No":code,"İş Kalemi":desc[:500],"Birim":unit,"Birim Fiyat":price,"Para Birimi":"TRY","Fiyat Yılı":int(period[:4]),"Kaynak":f"ÇŞİDB YFK — {cat}","Kaynak Dosya":src,"Kaynak Sayfa":pn,"Aktif":True,"Güncelleme Dönemi":period,"Yayın Tarihi":pub,"PDF URL":url,"PDF SHA256":dig,"Güven Skoru":conf,"İnceleme Durumu":"KONTROL BEKLİYOR"}

def yfk_pdf(blob,cat,grp,period,pub,vf,url,dig,src):
    try: import pdfplumber
    except: raise RuntimeError("pdfplumber yok.")
    if not blob.startswith(b"%PDF-"): raise ValueError("PDF değil.")
    rows=[]; seen=set()
    with pdfplumber.open(BytesIO(blob)) as pdf:
        for pn,page in enumerate(pdf.pages[:120],1):
            try:
                for tbl in page.extract_tables() or []:
                    for cells in tbl:
                        r=yfk_row(cells,pn,cat,grp,period,pub,vf,src,url,dig)
                        if r:
                            k=(r["Poz No"],norm(r["İş Kalemi"]),r["Birim"],r["Birim Fiyat"])
                            if k not in seen: rows.append(r); seen.add(k)
            except: pass
    return ef(pd.DataFrame(rows),YFK_CAND) if rows else ef(pd.DataFrame(),YFK_CAND)

def yfk_sync():
    info=yfk_discover(); cands=[]; logs=[]
    base=DATA_DIR/"yfk_pdfs"/info["period"]; base.mkdir(parents=True,exist_ok=True)
    old_c=load_ds("__GLOBAL__","yfk_candidates",YFK_CAND); old_r=load_ds("__GLOBAL__","yfk_runs",YFK_RUNS)
    for key,doc in info["docs"].items():
        rec={"_id":uid(),"Güncelleme Dönemi":info["period"],"Belge Kategorisi":doc["cat"],"PDF URL":doc["url"],"PDF SHA256":"","Aday Sayısı":0,"Durum":"BAŞLIYOR","Kontrol Tarihi":now_tr().isoformat(timespec="seconds"),"Not":""}
        try:
            r=yfk_req(doc["url"],pdf=True); blob=r.content; dg=hashlib.sha256(blob).hexdigest()
            rec["PDF SHA256"]=dg
            dup=old_r[(old_r["PDF SHA256"].astype(str)==dg)&(old_r["Durum"].astype(str).str.startswith("İNDİRİLDİ"))] if not old_r.empty else pd.DataFrame()
            if not dup.empty: rec["Durum"]="DEĞİŞMEDİ"; logs.append(rec); continue
            fn=f"{info['period']}_{re.sub(r'[^a-z0-9]+','_',norm(doc['cat']))}_{dg[:12]}.pdf"
            (base/fn).write_bytes(blob)
            fr=yfk_pdf(blob,doc["cat"],doc["group"],info["period"],info["pub"],info["vf"],doc["url"],dg,fn)
            if not fr.empty: cands.append(fr); rec["Aday Sayısı"]=len(fr); rec["Durum"]="İNDİRİLDİ — ADAYLAR KONTROL"
            else: rec["Durum"]="AYRIŞTIRMA YOK"
        except Exception as e: rec["Durum"]="HATA"; rec["Not"]=str(e)[:500]
        logs.append(rec)
    if cands: old_c=pd.concat([old_c,*cands],ignore_index=True).drop_duplicates(subset=["PDF SHA256","Kaynak Sayfa","Poz No","İş Kalemi","Birim","Birim Fiyat"],keep="last")
    old_r=pd.concat([old_r,pd.DataFrame(logs)],ignore_index=True)
    save_ds("__GLOBAL__","yfk_candidates",old_c,YFK_CAND); save_ds("__GLOBAL__","yfk_runs",old_r,YFK_RUNS)
    return info,logs,len(cands)

# DOSYA
ALLOWED={".dwg",".dxf",".pdf",".xlsx",".xlsm",".csv",".ifc",".png",".jpg",".jpeg",".webp"}
def validate_upload(up):
    name=Path(up.name).name; ext=Path(name).suffix.lower(); data=up.getvalue()
    if ext not in ALLOWED: raise ValueError("Uzantı.")
    if not data or len(data)>MAX_BYTES: raise ValueError(f"{MAX_MB} MB sınırı.")
    return name,ext,data,hashlib.sha256(data).hexdigest()

def read_dxf(path):
    try: import ezdxf
    except: raise RuntimeError("ezdxf yok.")
    doc=ezdxf.readfile(str(path)); msp=doc.modelspace(); code=int(doc.header.get("$INSUNITS",0) or 0)
    factor={1:.0254,2:.3048,4:.001,5:.01,6:1.0,7:1000.0}.get(code)
    detail=[]; layers={}
    for e in msp:
        k=e.dxftype(); layer=str(e.dxf.get("layer","0")); ln=None; ar=None
        try:
            if k=="LINE": ln=e.dxf.start.distance(e.dxf.end)
            elif k=="CIRCLE": r=float(e.dxf.radius); ln=2*math.pi*r; ar=math.pi*r*r
            elif k=="LWPOLYLINE": ln=float(e.length()); ar=float(e.area()) if e.closed else None
        except: pass
        detail.append({"Tip":k,"Katman":layer,"Uzunluk(m)":ln*factor if ln and factor else None,"Alan(m²)":ar*factor*factor if ar and factor else None})
        t=layers.setdefault(layer,{"Eleman":0,"Uzunluk":0.0,"Alan":0.0}); t["Eleman"]+=1
        if ln: t["Uzunluk"]+=ln
        if ar: t["Alan"]+=ar
    return {"factor":factor,"detail":detail,"layers":layers}

def excel_export(tables):
    out=BytesIO()
    with pd.ExcelWriter(out,engine="openpyxl") as w:
        for name,df in tables.items():
            safe=re.sub(r"[\[\]:*?/\\]","_",name)[:31] or "R"
            (df if isinstance(df,pd.DataFrame) else pd.DataFrame(df)).to_excel(w,sheet_name=safe,index=False)
        for ws in w.book.worksheets:
            ws.freeze_panes="A2"
            for c in ws[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="14263D")
    out.seek(0); return out.getvalue()

def fig_layout(fig,h=380):
    fig.update_layout(height=h,paper_bgcolor="white",plot_bgcolor="white",margin=dict(l=20,r=20,t=35,b=60),font=dict(family="Montserrat",size=11))
    fig.update_xaxes(showgrid=False); fig.update_yaxes(gridcolor="#E9EDF2"); return fig

# NAVİGASYON
SECTIONS={"Yönetim":["Dashboard","Performans"],"Proje":["Proje Künyesi","Dokümanlar","AI Belge"],"Metraj/Bütçe":["Excel Import","DXF Analiz","Birim Fiyat + YFK","Bütçe","Metraj Onay","Anomali"],"Maliyet":["Gerçekleşen","Fatura AI","EVM","Nakit Akışı","Monte Carlo EAC","Revizyon Diff","Kârlılık"],"Planlama":["İş Programı/CPM","Puantaj","Risk"],"Satın Alma":["Satın Alma","Stok"],"Sözleşme":["Sözleşme/Değişiklik","Taşeron Hakediş","İşveren Hakedişi","Onay Merkezi"],"AI":["AI Asistan","Fatura Öğrenme"],"Sistem":["Kullanıcı","Yedekleme","Excel Rapor","Sistem Bilgi"]}
ALL_PAGES=set(sum(SECTIONS.values(),[]))
ROLE_PAGES={"Yönetici":ALL_PAGES,"Proje Müdürü":ALL_PAGES-{"Kullanıcı","Sistem Bilgi","Yedekleme"},"Teknik Ofis":{"Dashboard","Performans","Proje Künyesi","Dokümanlar","AI Belge","Excel Import","DXF Analiz","Birim Fiyat + YFK","Bütçe","Metraj Onay","Anomali","İş Programı/CPM","EVM","Revizyon Diff","Onay Merkezi"},"Şantiye Şefi":{"Dashboard","Performans","Dokümanlar","DXF Analiz","Metraj Onay","İş Programı/CPM","Puantaj","Risk","Stok"},"Finans / Muhasebe":{"Dashboard","Performans","Gerçekleşen","Fatura AI","EVM","Nakit Akışı","Monte Carlo EAC","Kârlılık","Sözleşme/Değişiklik","Taşeron Hakediş","İşveren Hakedişi","Onay Merkezi"},"Satın Alma":{"Dashboard","Birim Fiyat + YFK","Satın Alma","Stok","Risk"},"Depo / Ambar":{"Dashboard","Stok"},"Görüntüleyici":{"Dashboard","Performans"}}

# ===== ANA UYGULAMA =====
if ENGINE is None: st.error(f"DB: {DB_ERR}"); st.stop()
try: init_db()
except Exception as e: st.error(f"DB init: {e}"); st.stop()
auth_gate(); USER=st.session_state.auth_user

projects=load_ds("__GLOBAL__","project",PROJECT)
pnames=sorted(set(projects["Proje Adı"].dropna().astype(str).str.strip())-{""}) if not projects.empty else []
if "current_project" not in st.session_state: st.session_state.current_project=pnames[0] if pnames else "Yeni Proje"

st.markdown(f'<div class="ea-head"><h1>🏗️ {APP_NAME}</h1><small>V{APP_VERSION} · {USER["display_name"]} · {USER["role"]} · {now_tr():%d.%m.%Y %H:%M}</small></div>',unsafe_allow_html=True)
with st.sidebar:
    st.markdown("## EMİR ASEL")
    st.caption(f"{USER['display_name']}")
    opts=list(dict.fromkeys(pnames+["Yeni Proje"]))
    if st.session_state.current_project not in opts: st.session_state.current_project=opts[0]
    chosen=st.selectbox("Aktif Proje",opts,index=opts.index(st.session_state.current_project))
    if chosen!=st.session_state.current_project: st.session_state.current_project=chosen; st.rerun()
    visible_sections=[s for s,pgs in SECTIONS.items() if any(p in ROLE_PAGES[USER["role"]] for p in pgs)]
    sec=st.selectbox("BÖLÜM",visible_sections,key="nav_sec")
    pages=[p for p in SECTIONS[sec] if p in ROLE_PAGES[USER["role"]]]
    page=st.radio("İŞLEM",pages,key=f"nav_{sec}")
    st.divider(); st.caption(f"DB: {DB_KIND}")
    if st.button("Çıkış",width="stretch"): st.session_state.pop("auth_user",None); st.rerun()

PROJECT_KEY=st.session_state.current_project if st.session_state.current_project!="Yeni Proje" else "__DRAFT_PROJECT__"
meta=projects[projects["Proje Adı"].astype(str)==st.session_state.current_project] if not projects.empty else pd.DataFrame()
META=meta.iloc[0].to_dict() if not meta.empty else {}
BASE_CURRENCY=cur_code(META.get("Para Birimi"),"TRY")
REPORT_DATE=st.date_input("Rapor Tarihi",now_tr().date(),key=f"rd_{PROJECT_KEY}")

DATA={k:load_ds("__GLOBAL__" if k in GLOBAL_DATASETS else PROJECT_KEY,k,v) for k,v in DS.items() if k!="project"}
PLAN_DF=calc_plan(DATA["plan"],BASE_CURRENCY); ACTUAL_DF=calc_actual(DATA["actual"],BASE_CURRENCY)
EVM=evm_calc(PLAN_DF,ACTUAL_DF,DATA["schedule"],REPORT_DATE,BASE_CURRENCY)
SUMMARY=summary_costs(PLAN_DF,ACTUAL_DF,BASE_CURRENCY); PROGRESS=wprogress(DATA["schedule"])
CPM=cpm_calc(DATA["schedule"])
MEM=DATA["memory"]

def save_editor(ds_name,edited,cols,recalc=None):
    frame=ef(edited,cols)
    if recalc=="plan": frame=calc_plan(frame,BASE_CURRENCY)
    elif recalc=="actual": frame=calc_actual(frame,BASE_CURRENCY)
    elif recalc=="risk": frame["Risk Skoru"]=pd.to_numeric(frame["Olasılık (1-5)"],errors="coerce").fillna(0)*pd.to_numeric(frame["Etki (1-5)"],errors="coerce").fillna(0)
    elif recalc=="emp_hak": frame=emp_hak_calc(frame)
    save_ds("__GLOBAL__" if ds_name in GLOBAL_DATASETS else PROJECT_KEY,ds_name,frame,cols)
    st.success(f"{len(frame)} kayıt."); st.rerun()

# SAYFA YÖNLENDİRİCİ
if page=="Dashboard":
    st.title("Dashboard")
    c1,c2,c3,c4=st.columns(4)
    c1.metric("Bütçe",money(EVM["BAC"],BASE_CURRENCY))
    c2.metric("Gerçekleşen",money(EVM["AC"],BASE_CURRENCY))
    c3.metric("İlerleme",pct_fmt(PROGRESS))
    c4.metric("EAC",money(EVM["EAC"],BASE_CURRENCY) if pd.notna(EVM["EAC"]) else "—")
    c5,c6,c7=st.columns(3)
    c5.metric("CPI",f'{EVM["CPI"]:.2f}' if pd.notna(EVM["CPI"]) else "—")
    c6.metric("SPI",f'{EVM["SPI"]:.2f}' if pd.notna(EVM["SPI"]) else "—")
    c7.metric("Risk",int(DATA["risk"]["Durum"].map(lambda x:norm(x) in{"acik","open"}).sum()) if not DATA["risk"].empty else 0)
    if not SUMMARY.empty:
        fig=px.bar(SUMMARY.melt(id_vars="Ana Kategori",value_vars=["Bütçe","Gerçekleşen"],var_name="Tür",value_name="Tutar"),x="Ana Kategori",y="Tutar",color="Tür",barmode="group",color_discrete_sequence=[NAVY,RED])
        fig.update_layout(xaxis_tickangle=-35,margin=dict(b=130))
        st.plotly_chart(fig_layout(fig,430),width="stretch")

elif page=="Performans":
    st.title("Proje Performansı")
    st.dataframe(pd.DataFrame([{"Gösterge":k,"Değer":v} for k,v in EVM.items()]),width="stretch",hide_index=True)

elif page=="Proje Künyesi":
    st.title("Proje Künyesi")
    with st.form("proj"):
        a,b=st.columns(2); name=a.text_input("Proje Adı",META.get("Proje Adı","")); owner=b.text_input("İşveren",META.get("İşveren",""))
        c,d=st.columns(2); loc=c.text_input("Lokasyon",META.get("Lokasyon","")); cno=d.text_input("Sözleşme No",META.get("Sözleşme No",""))
        e,f=st.columns(2); start=e.date_input("Başlangıç",now_tr().date()); end=f.date_input("Bitiş",now_tr().date()+timedelta(days=365))
        g,h=st.columns(2); cur=g.selectbox("Para Birimi",CURRENCIES,index=CURRENCIES.index(BASE_CURRENCY)); vat=h.number_input("KDV %",0.0,100.0,20.0)
        sub=st.form_submit_button("Kaydet",type="primary")
    if sub and name.strip():
        rec={"_id":str(META.get("_id") or uid()),"Proje Adı":name.strip(),"İşveren":owner,"Lokasyon":loc,"Sözleşme No":cno,"Başlangıç Tarihi":start,"Bitiş Tarihi":end,"Para Birimi":cur,"KDV Oranı":vat/100}
        rest=projects[projects["_id"].astype(str)!=rec["_id"]] if not projects.empty else pd.DataFrame(columns=PROJECT)
        save_ds("__GLOBAL__","project",pd.concat([rest,pd.DataFrame([rec])],ignore_index=True),PROJECT)
        st.session_state.current_project=name.strip(); st.rerun()

elif page=="Dokümanlar":
    st.title("Doküman Merkezi")
    up=st.file_uploader("Dosya",type=["pdf","xlsx","dxf","dwg","png","jpg"])
    if up and st.button("Yükle",type="primary"):
        try:
            name,ext,data,dig=validate_upload(up)
            with ENGINE.begin() as con:
                con.execute(text("INSERT INTO ea_file_assets(id,project_key,original_name,stored_name,extension,sha256,size_bytes,uploaded_by,created_at) VALUES(:i,:p,:n,:s,:e,:h,:z,:u,:t)"),{"i":uid(),"p":PROJECT_KEY,"n":name,"s":uid()+ext,"e":ext,"h":dig,"z":len(data),"u":USER["username"],"t":now_tr().isoformat()})
            st.success(f"Yüklendi: {name}")
        except Exception as e: st.error(str(e))
    with ENGINE.connect() as con: rows=con.execute(text("SELECT original_name,extension,size_bytes,uploaded_by,created_at FROM ea_file_assets WHERE project_key=:p"),{"p":PROJECT_KEY}).fetchall()
    if rows: st.dataframe(pd.DataFrame([dict(r._mapping) for r in rows]),width="stretch",hide_index=True)

elif page=="AI Belge":
    st.title("AI Belge Analizi")
    up=st.file_uploader("PDF/Görsel",type=["pdf","png","jpg","jpeg"])
    ctx=st.text_area("Bağlam")
    if up and st.button("Analiz Et",type="primary"):
        try:
            ext=Path(up.name).suffix.lower()
            if ext==".pdf":
                try:
                    import pdfplumber
                    with pdfplumber.open(BytesIO(up.getvalue())) as pdf: txt="\n".join((p.extract_text() or "") for p in pdf.pages[:80])
                except: txt=""
                if len(txt.strip())<50: st.warning("PDF metinsiz. Görsel analiz için PNG/JPG kullan.")
                else: st.text_area("AI Sonucu",ai_text(txt,ctx),height=400)
            else: st.text_area("AI Sonucu",ai_image(up.getvalue(),mimetypes.guess_type(up.name)[0] or "image/png",ctx),height=400)
        except Exception as e: st.error(str(e))

elif page=="Excel Import":
    st.title("Excel Metraj Import")
    up=st.file_uploader("Excel",type=["xlsx","xlsm"])
    if up:
        try: sheets=pd.read_excel(up,sheet_name=None)
        except Exception as e: st.error(f"Excel: {e}"); st.stop()
        sh=st.selectbox("Sayfa",list(sheets)); raw=sheets[sh]
        st.dataframe(raw.head(10),width="stretch",hide_index=True)
        cols=["İş Kalemi","Birim","Miktar","Poz No"]; mp={}
        for c in cols:
            cands={"İş Kalemi":["İş Kalemi","Açıklama","Tanım"],"Birim":["Birim"],"Miktar":["Miktar","Metraj"],"Poz No":["Poz No","Poz","Kod"]}[c]
            g=find_col(raw,cands); opts=["—"]+list(raw.columns); mp[c]=st.selectbox(c,opts,index=opts.index(g) if g in opts else 0,key=f"mp_{c}")
        if st.button("Taslak Bütçeye Aktar",type="primary"):
            if mp["İş Kalemi"]=="—" or mp["Miktar"]=="—": st.error("Zorunlu sütunlar."); st.stop()
            rows=[]
            for _,r in raw.iterrows():
                d=str(r[mp["İş Kalemi"]] or "").strip()
                if not d: continue
                q=num(r[mp["Miktar"]]); 
                if q<=0: continue
                rows.append({"_id":uid(),"WBS":f"IMP-{uid()[:6]}","Poz No":str(r[mp["Poz No"]] or ""),"İş Kalemi":d,"Ana Kategori":categorize(d),"Birim":str(r[mp["Birim"]] or "") if mp["Birim"]!="—" else "adet","Miktar":q,"Birim Fiyat":0,"Para Birimi":BASE_CURRENCY,"Kur":1.0,"Metraj Durumu":"Kontrol Bekliyor","Açıklama":f"Excel: {up.name}"})
            if not rows: st.error("Geçerli satır yok."); st.stop()
            newp=calc_plan(pd.DataFrame(rows),BASE_CURRENCY)
            save_ds(PROJECT_KEY,"plan",pd.concat([PLAN_DF,newp],ignore_index=True),PLAN)
            st.success(f"{len(rows)} satır aktarıldı."); st.rerun()

elif page=="DXF Analiz":
    st.title("DXF Analizi")
    up=st.file_uploader("DXF",type=["dxf"])
    if up:
        try:
            name,ext,data,dig=validate_upload(up)
            with tempfile.TemporaryDirectory() as td:
                p=Path(td)/name; p.write_bytes(data); res=read_dxf(p)
            st.metric("Birim Faktörü",res["factor"] or "Yok")
            if res["layers"]: st.dataframe(pd.DataFrame([{"Katman":k,"Eleman":v["Eleman"],"Uzunluk":v["Uzunluk"],"Alan":v["Alan"]} for k,v in res["layers"].items()]),width="stretch",hide_index=True)
        except Exception as e: st.error(str(e))

elif page=="Birim Fiyat + YFK":
    st.title("Birim Fiyat Kütüphanesi + YFK")
    st.markdown(f"**Resmi Kaynak:** [{YFK_INDEX}]({YFK_INDEX})")
    runs=DATA["yfk_runs"]
    with st.expander("Resmî YFK Otomatik Güncelleme",expanded=True):
        if st.button("YFK'yı Şimdi Kontrol Et",type="primary"):
            with st.spinner("İndiriliyor…"):
                try:
                    info,logs,nc=yfk_sync(); st.success(f"{info['period']} · {nc} aday PDF"); st.dataframe(pd.DataFrame(logs),width="stretch",hide_index=True); st.rerun()
                except Exception as e: st.error(f"YFK: {e}")
    pending=DATA["yfk_candidates"][DATA["yfk_candidates"]["İnceleme Durumu"].astype(str).eq("KONTROL BEKLİYOR")] if not DATA["yfk_candidates"].empty else pd.DataFrame()
    if not pending.empty:
        st.subheader(f"YFK Adayları ({len(pending)})")
        ce=st.data_editor(pending,num_rows="dynamic",width="stretch",hide_index=True,key="yfk_ce")
        if st.button("İşaretlileri Aktar",type="primary"):
            acc=ce[ce["Kabul"].fillna(False)]
            if not acc.empty:
                lib=DATA["library"].copy(); ray=DATA["rayic"].copy(); upd=DATA["yfk_candidates"].copy()
                for _,r in acc.iterrows():
                    if str(r.get("Kayıt Grubu"))=="Rayiç":
                        ray=pd.concat([ray,pd.DataFrame([{"_id":uid(),"Rayiç No":r.get("Poz No"),"Açıklama":r.get("İş Kalemi"),"Birim":r.get("Birim"),"Rayiç Fiyat":r.get("Birim Fiyat"),"Para Birimi":"TRY","Fiyat Yılı":r.get("Fiyat Yılı"),"Kaynak Türü":r.get("Belge Kategorisi"),"Durum":"ONAYLANDI"}])],ignore_index=True)
                    else:
                        lib=pd.concat([lib,pd.DataFrame([{"_id":uid(),"Poz No":r.get("Poz No"),"İş Kalemi":r.get("İş Kalemi"),"Birim":r.get("Birim"),"Birim Fiyat":r.get("Birim Fiyat"),"Para Birimi":"TRY","Fiyat Yılı":r.get("Fiyat Yılı"),"Kaynak":r.get("Kaynak"),"Aktif":True,"Güncelleme Dönemi":r.get("Güncelleme Dönemi")}])],ignore_index=True)
                    upd.loc[upd["_id"].astype(str)==str(r["_id"]),"İnceleme Durumu"]="AKTARILDI"
                save_ds("__GLOBAL__","library",ef(lib,LIB),LIB); save_ds("__GLOBAL__","rayic",ef(ray,RAYIC),RAYIC); save_ds("__GLOBAL__","yfk_candidates",ef(upd,YFK_CAND),YFK_CAND)
                st.success(f"{len(acc)} aktarıldı."); st.rerun()
    if not runs.empty: st.subheader("YFK Geçmişi"); st.dataframe(runs.sort_values("Kontrol Tarihi",ascending=False),width="stretch",hide_index=True)
    if not DATA["library"].empty: st.subheader(f"Şirket Kütüphanesi ({len(DATA['library'])})"); st.dataframe(DATA["library"].drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True,height=350)

elif page=="Bütçe":
    st.title("Bütçe ve Metraj")
    ed=st.data_editor(PLAN_DF,num_rows="dynamic",width="stretch",hide_index=True,column_config={"_id":st.column_config.Column(hidden=True),"Ana Kategori":st.column_config.SelectboxColumn(options=CATEGORIES),"Metraj Durumu":st.column_config.TextColumn(disabled=True),"Planlanan Tutar":st.column_config.NumberColumn(disabled=True)},key=f"bud_{PROJECT_KEY}")
    if st.button("Kaydet",type="primary"): save_editor("plan",ed,PLAN,"plan")
    st.metric("Resmi BAC",money(EVM["BAC"],BASE_CURRENCY))

elif page=="Metraj Onay":
    st.title("Metraj Onayları")
    st.dataframe(DATA["plan"].drop(columns=["_id"],errors="ignore") if not DATA["plan"].empty else pd.DataFrame(),width="stretch",hide_index=True)

elif page=="Anomali":
    st.title("🔍 Anomali Tespiti")
    st.caption("IQR yöntemiyle aykırı değerleri yakalar.")
    for label,df in [("Bütçe",PLAN_DF),("Gerçekleşen",ACTUAL_DF)]:
        st.subheader(label)
        a=detect_anomalies(df)
        if a.empty: st.success("Aykırı değer yok.")
        else: st.warning(f"{len(a)} aykırı değer"); st.dataframe(a,width="stretch",hide_index=True)

elif page=="Gerçekleşen":
    st.title("Gerçekleşen Maliyet")
    ed=st.data_editor(ACTUAL_DF,num_rows="dynamic",width="stretch",hide_index=True,column_config={"_id":st.column_config.Column(hidden=True),"Gerçekleşen Tutar":st.column_config.NumberColumn(disabled=True),"Durum":st.column_config.TextColumn(disabled=True)},key=f"act_{PROJECT_KEY}")
    if st.button("Kaydet",type="primary"): save_editor("actual",ed,ACTUAL,"actual")

elif page=="Fatura AI":
    st.title("📄 Fatura AI")
    up=st.file_uploader("Fatura PDF/Excel",type=["pdf","xlsx","xlsm"])
    mem=DATA["memory"]
    if mem.empty: st.info("Henüz öğrenilmiş eşleştirme yok. Fatura işleyince sistem öğrenmeye başlayacak.")
    else: st.caption(f"📚 {len(mem)} öğrenilmiş eşleştirme hazır")
    if up:
        try:
            if up.name.lower().endswith(".pdf"):
                try:
                    import pdfplumber
                    with pdfplumber.open(BytesIO(up.getvalue())) as pdf: txt="\n".join((p.extract_text() or "") for p in pdf.pages[:10])
                except: txt=""
                if len(txt.strip())<50: st.warning("PDF metinsiz. Excel kullanın veya OCR ekleyin.")
                else:
                    if st.button("AI ile Satır Çıkar",type="primary"):
                        try:
                            c=ai_client(); r=c.chat.completions.create(model=os.getenv("OPENAI_MODEL") or "gpt-4o-mini",messages=[{"role":"system","content":"JSON döndür: invoice_no, invoice_date, supplier, currency, lines[].description, quantity, unit, unit_price, amount. Türkçe metinleri doğru oku."},{"role":"user","content":txt[:30000]}],max_tokens=1500)
                            js=r.choices[0].message.content; js=re.sub(r"^```(?:json)?\s*|\s*```$","",js,flags=re.I)
                            obj=json.loads(js)
                            rows=[]
                            for ln in obj.get("lines",[]):
                                d=str(ln.get("description","")).strip()
                                if not d: continue
                                sug=suggest_from_memory(obj.get("supplier",""),d,mem)
                                rows.append({"_id":uid(),"WBS":"","Poz No":str(ln.get("poz","")),"İş Kalemi":d,"Ana Kategori":(sug or {}).get("Ana Kategori",categorize(d)),"Birim":str(ln.get("unit","")),"Gerçekleşen Miktar":num(ln.get("quantity")),"Gerçekleşen Birim Fiyat":num(ln.get("unit_price")),"Gerçekleşen Tutar":0,"Para Birimi":cur_code(ln.get("currency") or obj.get("currency"),BASE_CURRENCY),"Kur":1.0,"Tarih":obj.get("invoice_date"),"Tedarikçi":obj.get("supplier",""),"Belge No":obj.get("invoice_no",""),"Kayıt Türü":"Fatura","Durum":"Kontrol Bekliyor","Öneri":(sug or {}).get("Ana Kategori","Yok")})
                            if rows:
                                st.dataframe(pd.DataFrame(rows).drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)
                                st.session_state["fatura_cand"]=rows; st.success(f"{len(rows)} satır çıkarıldı. Öneriler Fatura Öğrenme hafızasından.")
                        except Exception as e: st.error(f"AI: {e}")
            if "fatura_cand" in st.session_state and st.button("Onayla ve Kaydet",type="primary"):
                try:
                    for r in st.session_state["fatura_cand"]:
                        mem=learn_mapping(PROJECT_KEY,r["Tedarikçi"],r["İş Kalemi"],r["Ana Kategori"],r.get("Poz No",""),mem)
                    new_df=pd.concat([DATA["actual"],pd.DataFrame(st.session_state["fatura_cand"])],ignore_index=True)
                    save_ds(PROJECT_KEY,"actual",calc_actual(new_df,BASE_CURRENCY),ACTUAL)
                    save_ds(PROJECT_KEY,"memory",mem,MEMORY)
                    st.session_state.pop("fatura_cand",None)
                    st.success(f"Kaydedildi ve {len(st.session_state.get('fatura_cand',[]))} eşleştirme öğrenildi."); st.rerun()
                except Exception as e: st.error(str(e))
        except Exception as e: st.error(f"Fatura: {e}")

elif page=="EVM":
    st.title("EVM Performans")
    c=st.columns(5)
    for col,k in zip(c,["BAC","PV","EV","AC","EAC"]): col.metric(k,money(EVM[k],BASE_CURRENCY) if pd.notna(EVM[k]) else "—")
    c2=st.columns(4)
    c2[0].metric("CPI",f'{EVM["CPI"]:.2f}' if pd.notna(EVM["CPI"]) else "—")
    c2[1].metric("SPI",f'{EVM["SPI"]:.2f}' if pd.notna(EVM["SPI"]) else "—")
    c2[2].metric("CV",money(EVM["CV"],BASE_CURRENCY))
    c2[3].metric("SV",money(EVM["SV"],BASE_CURRENCY))

elif page=="Nakit Akışı":
    st.title("Nakit Akışı")
    ed=st.data_editor(DATA["cash"],num_rows="dynamic",width="stretch",hide_index=True,column_config={"_id":st.column_config.Column(hidden=True)},key=f"csh_{PROJECT_KEY}")
    if st.button("Kaydet",type="primary"): save_editor("cash",ed,CASH)

elif page=="Monte Carlo EAC":
    st.title("🎲 Monte Carlo EAC")
    st.caption("1000 simülasyon — bitirme maliyeti P10/P50/P90.")
    c1,c2,c3=st.columns(3)
    c1.metric("BAC",money(EVM["BAC"],BASE_CURRENCY))
    c2.metric("AC",money(EVM["AC"],BASE_CURRENCY))
    c3.metric("EV",money(EVM["EV"],BASE_CURRENCY))
    if st.button("Simülasyonu Başlat",type="primary"):
        r=monte_carlo_eac(EVM["BAC"],EVM["AC"],EVM["EV"],DATA["risk"])
        if r is None: st.error("BAC > 0 gerekli.")
        else:
            c1,c2,c3,c4=st.columns(4)
            c1.metric("P10 (İyi)",money(r["P10"],BASE_CURRENCY))
            c2.metric("P50 (Beklenen)",money(r["P50"],BASE_CURRENCY))
            c3.metric("P90 (Kötü)",money(r["P90"],BASE_CURRENCY))
            c4.metric("Ortalama",money(r["Ort"],BASE_CURRENCY))
            fig=px.histogram(x=r["Sims"],nbins=50,labels={"x":"EAC"})
            for v,lbl,cl in [(r["P10"],"P10","green"),(r["P50"],"P50","blue"),(r["P90"],"P90","red"),(EVM["BAC"],"BAC","black")]: fig.add_vline(x=v,line_dash="dot",line_color=cl,annotation_text=lbl)
            st.plotly_chart(fig_layout(fig,400),width="stretch")

elif page=="Revizyon Diff":
    st.title("📊 Revizyon Karşılaştırma")
    c1,c2=st.columns(2)
    with c1: a_up=st.file_uploader("Revizyon A (Eski)",type=["xlsx","xlsm"],key="revA")
    with c2: b_up=st.file_uploader("Revizyon B (Yeni)",type=["xlsx","xlsm"],key="revB")
    if a_up and b_up and st.button("Karşılaştır",type="primary"):
        try:
            A=pd.read_excel(a_up,sheet_name=0); B=pd.read_excel(b_up,sheet_name=0)
            kc=[]; vc=[]
            for tgt in [["Poz No","Poz"],["İş Kalemi","Açıklama"]]:
                c=find_col(A,tgt)
                if c: kc.append(c)
            for tgt in [["Miktar","Metraj"],["Birim Fiyat","BF","Fiyat"]]:
                c=find_col(A,tgt)
                if c: vc.append(c)
            if not kc or not vc: st.error("Sütunlar bulunamadı."); st.stop()
            diff=compare_revisions(A,B,kc,vc)
            if diff.empty: st.success("Fark yok.")
            else:
                tot=float(pd.to_numeric(diff.get("Tutar Farkı",0),errors="coerce").fillna(0).sum())
                c1,c2,c3,c4=st.columns(4)
                c1.metric("Toplam",len(diff)); c2.metric("Eklendi",int((diff["Durum"]=="EKLENDİ").sum())); c3.metric("Silindi",int((diff["Durum"]=="SİLİNDİ").sum())); c4.metric("Maliyet",money(tot,BASE_CURRENCY))
                st.dataframe(diff,width="stretch",hide_index=True)
                buf=BytesIO(); diff.to_excel(buf,index=False,engine="openpyxl")
                st.download_button("Excel İndir",buf.getvalue(),f"rev_diff_{now_tr():%Y%m%d}.xlsx")
        except Exception as e: st.error(f"Hata: {e}")

elif page=="Kârlılık":
    st.title("Proje Kârlılığı")
    bac=EVM["BAC"]; ac=EVM["AC"]; ev=EVM["EV"]; cpi=EVM["CPI"]
    eac=bac/cpi if pd.notna(cpi) and cpi>0 else np.nan
    st.metric("BAC",money(bac,BASE_CURRENCY))
    st.metric("EAC",money(eac,BASE_CURRENCY) if pd.notna(eac) else "—")
    st.metric("VAC",money(bac-eac,BASE_CURRENCY) if pd.notna(eac) else "—")

elif page=="İş Programı/CPM":
    st.title("İş Programı / CPM")
    ed=st.data_editor(DATA["schedule"],num_rows="dynamic",width="stretch",hide_index=True,column_config={"_id":st.column_config.Column(hidden=True)},key=f"sch_{PROJECT_KEY}")
    if st.button("Kaydet",type="primary"): save_editor("schedule",ed,SCHEDULE)
    if not CPM.empty: st.dataframe(CPM.drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page=="Puantaj":
    st.title("İşçilik ve Puantaj")
    if not DATA["labor"].empty: st.dataframe(DATA["labor"].drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page=="Risk":
    st.title("Risk Yönetimi")
    ed=st.data_editor(DATA["risk"],num_rows="dynamic",width="stretch",hide_index=True,column_config={"_id":st.column_config.Column(hidden=True)},key=f"rsk_{PROJECT_KEY}")
    if st.button("Kaydet",type="primary"): save_editor("risk",ed,RISK,"risk")

elif page=="Satın Alma":
    st.title("Satın Alma")
    if not DATA["purchase"].empty: st.dataframe(DATA["purchase"].drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page=="Stok":
    st.title("Depo / Stok")
    if not DATA["inventory"].empty: st.dataframe(DATA["inventory"].drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page=="Sözleşme/Değişiklik":
    st.title("Sözleşme ve Değişiklik")
    if not DATA["contracts"].empty: st.dataframe(DATA["contracts"].drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)
    if not DATA["changes"].empty: st.dataframe(DATA["changes"].drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page=="Taşeron Hakediş":
    st.title("Taşeron Hakediş")
    if not DATA["hakediş"].empty: st.dataframe(DATA["hakediş"].drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page=="İşveren Hakedişi":
    st.title("İşveren Hakedişi")
    e_h=emp_hak_calc(DATA["employer_hakediş"])
    if not e_h.empty:
        c1,c2,c3,c4=st.columns(4)
        c1.metric("Brüt",money(e_h["Brüt İmalat Tutarı"].sum(),BASE_CURRENCY))
        c2.metric("Hakediş",money(e_h["Hakediş Toplamı"].sum(),BASE_CURRENCY))
        c3.metric("Tahsilat",money(e_h["Tahsil Edilen"].sum(),BASE_CURRENCY))
        c4.metric("Kalan",money(e_h["Kalan Alacak"].sum(),BASE_CURRENCY))
        st.dataframe(e_h.drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)

elif page=="Onay Merkezi":
    st.title("Onay Merkezi")
    if not DATA["approvals"].empty:
        st.dataframe(DATA["approvals"].drop(columns=["_id"],errors="ignore"),width="stretch",hide_index=True)
        pend=DATA["approvals"][DATA["approvals"]["Durum"].astype(str)=="Bekliyor"]
        if not pend.empty and USER["role"] in {"Yönetici","Proje Müdürü"}:
            opts={f"{r['Talep Başlığı']} · {money(r['Tutar'],BASE_CURRENCY)}":str(r["_id"]) for _,r in pend.iterrows()}
            sel=st.selectbox("Seç",list(opts)); note=st.text_area("Not")
            c1,c2=st.columns(2)
            if c1.button("Onayla",type="primary"):
                ap=DATA["approvals"].copy(); ap.loc[ap["_id"].astype(str)==opts[sel],"Durum"]="Onaylandı"; ap.loc[ap["_id"].astype(str)==opts[sel],"Karar Veren"]=USER["username"]; ap.loc[ap["_id"].astype(str)==opts[sel],"Karar Notu"]=note
                save_ds(PROJECT_KEY,"approvals",ap,APPROVAL); st.rerun()
            if c2.button("Reddet"):
                ap=DATA["approvals"].copy(); ap.loc[ap["_id"].astype(str)==opts[sel],"Durum"]="Reddedildi"; ap.loc[ap["_id"].astype(str)==opts[sel],"Karar Veren"]=USER["username"]; ap.loc[ap["_id"].astype(str)==opts[sel],"Karar Notu"]=note
                save_ds(PROJECT_KEY,"approvals",ap,APPROVAL); st.rerun()

elif page=="AI Asistan":
    st.title("🤖 AI Proje Asistanı")
    issues=[]
    if EVM["BAC"]>0 and EVM["AC"]>EVM["BAC"]*1.10: issues.append(("Kritik","Maliyet",f"AC bütçenin %{((EVM['AC']/EVM['BAC'])-1)*100:.0f} üstünde"))
    if pd.notna(EVM["CPI"]) and EVM["CPI"]<0.90: issues.append(("Kritik","CPI",f"CPI {EVM['CPI']:.2f}"))
    if pd.notna(EVM["SPI"]) and EVM["SPI"]<0.90: issues.append(("Kritik","SPI",f"SPI {EVM['SPI']:.2f}"))
    if not DATA["risk"].empty:
        hr=(pd.to_numeric(DATA["risk"].get("Risk Skoru",0),errors="coerce").fillna(0)>=15).sum()
        if hr>0: issues.append(("Kritik","Risk",f"{hr} yüksek risk"))
    if not DATA["approvals"].empty:
        pend=len(DATA["approvals"][DATA["approvals"]["Durum"].astype(str)=="Bekliyor"])
        if pend>0: issues.append(("Bilgi","Onay",f"{pend} onay bekliyor"))
    if not issues: st.success("Kritik bulgu yok.")
    else:
        for s,a,t in issues:
            col={"Kritik":"#C83E4D","Uyarı":"#D98E04","Bilgi":"#457B9D"}.get(s,"#666")
            st.markdown(f"<div style='border-left:4px solid {col};padding:8px 12px;background:#F8FAFC;margin:4px 0;border-radius:6px'><b>{a}:</b> {t}</div>",unsafe_allow_html=True)
    st.divider()
    if "ai_chat" not in st.session_state: st.session_state.ai_chat=[]
    for m in st.session_state.ai_chat:
        with st.chat_message(m["role"]): st.write(m["content"])
    q=st.chat_input("Sor: EAC, anomali, özet, risk...")
    if q:
        st.session_state.ai_chat.append({"role":"user","content":q})
        try:
            qn=norm(q)
            if "eac" in qn or "bitirme" in qn:
                r=monte_carlo_eac(EVM["BAC"],EVM["AC"],EVM["EV"],DATA["risk"])
                ans=f"**EAC:**\n- P10: {money(r['P10'],BASE_CURRENCY)}\n- P50: {money(r['P50'],BASE_CURRENCY)}\n- P90: {money(r['P90'],BASE_CURRENCY)}" if r else "Veri yetersiz."
            elif "anomali" in qn or "aykırı" in qn:
                a=detect_anomalies(PLAN_DF); ans=f"**{len(a)} anomali bulundu.**" if not a.empty else "Anomali yok."
            elif "özet" in qn:
                ans=f"**Özet:**\n- BAC: {money(EVM['BAC'],BASE_CURRENCY)}\n- AC: {money(EVM['AC'],BASE_CURRENCY)}\n- CPI: {EVM['CPI']:.2f}" if pd.notna(EVM['CPI']) else "Özet: BAC var, CPI yok."
            else:
                c=ai_client(); r=c.chat.completions.create(model=os.getenv("OPENAI_MODEL") or "gpt-4o-mini",messages=[{"role":"system","content":f"İnşaat asistanı. Veri: BAC={EVM['BAC']}, AC={EVM['AC']}, CPI={EVM['CPI']}, SPI={EVM['SPI']}. Sayı uydurma."},{"role":"user","content":q}],max_tokens=500)
                ans=r.choices[0].message.content
            st.session_state.ai_chat.append({"role":"assistant","content":ans}); st.rerun()
        except Exception as e: st.error(f"AI: {e}")

elif page=="Fatura Öğrenme":
    st.title("🧠 Fatura Öğrenme")
    mem=DATA["memory"]
    if mem.empty: st.info("Henüz öğrenilmiş eşleştirme yok.")
    else:
        c1,c2=st.columns(2); c1.metric("Eşleştirme",len(mem)); c2.metric("Toplam Kullanım",int(pd.to_numeric(mem["Kullanım"],errors="coerce").fillna(0).sum()))
        st.dataframe(mem.sort_values("Kullanım",ascending=False),width="stretch",hide_index=True)
        if st.button("Hafızayı Temizle",type="secondary"): save_ds(PROJECT_KEY,"memory",pd.DataFrame(columns=MEMORY),MEMORY); st.rerun()

elif page=="Kullanıcı":
    st.title("Kullanıcı Yönetimi")
    if USER["role"]!="Yönetici": st.error("Yönetici gerekli."); st.stop()
    with st.form("nu"):
        a,b=st.columns(2); un=a.text_input("Kullanıcı adı"); dn=b.text_input("Ad Soyad")
        c,d=st.columns(2); role=c.selectbox("Rol",ROLES[1:]); pw=d.text_input("Parola (12+)",type="password")
        if st.form_submit_button("Oluştur"):
            try: create_user(un,dn,role,pw); st.success("Oluşturuldu."); st.rerun()
            except Exception as e: st.error(str(e))
    st.dataframe(list_users(),width="stretch",hide_index=True)

elif page=="Yedekleme":
    st.title("Yedekleme / Denetim")
    bkp={"app":APP_NAME,"project":PROJECT_KEY,"data":{k:v.to_dict("records") for k,v in DATA.items()}}
    st.download_button("JSON Yedek",pl(bkp).encode(),f"backup_{now_tr():%Y%m%d}.json","application/json")

elif page=="Excel Rapor":
    st.title("Excel Raporları")
    tabs={"Bütçe":PLAN_DF.drop(columns=["_id"],errors="ignore"),"Gerçekleşen":ACTUAL_DF.drop(columns=["_id"],errors="ignore"),"Riskler":DATA["risk"].drop(columns=["_id"],errors="ignore"),"EVM":pd.DataFrame([EVM])}
    try:
        x=excel_export(tabs); st.download_button("Excel İndir",x,f"rapor_{now_tr():%Y%m%d}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
    except Exception as e: st.error(f"Excel: {e}")

elif page=="Sistem Bilgi":
    st.title("Sistem Bilgileri")
    st.write({"Sürüm":APP_VERSION,"DB":DB_KIND,"Proje":PROJECT_KEY,"Para":BASE_CURRENCY,"OpenAI":"✅" if os.getenv("OPENAI_API_KEY") else "❌","Token":"✅" if os.getenv("EA_BOOTSTRAP_TOKEN") else "❌"})

else: st.info(f"'{page}' modülü hazırlanıyor.")

st.divider(); st.caption(f"{APP_NAME} · V{APP_VERSION} · {DB_KIND.upper()} · AI karar destek amaçlıdır.")
