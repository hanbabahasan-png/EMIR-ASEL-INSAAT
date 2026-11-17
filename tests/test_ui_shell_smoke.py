import os, sys, types, importlib.util, tempfile, hashlib, json, ast
from pathlib import Path
from sqlalchemy import create_engine, text

APP_DIR=Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path: sys.path.insert(0,str(APP_DIR))
SRC=(APP_DIR/'app.py').read_text(encoding='utf-8')
TREE=ast.parse(SRC)
SECTIONS=None
for n in TREE.body:
    if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='SECTIONS' for t in n.targets):
        SECTIONS=ast.literal_eval(n.value); break
if SECTIONS is None: raise SystemExit('SECTIONS not found')

class Stop(BaseException): pass
class Cfg:
    def __getattr__(self,n): return lambda *a,**k: object()
class State(dict):
    __getattr__=dict.get
    def __setattr__(self,k,v): self[k]=v
class UI:
    def __init__(self): self.session_state=State(); self.secrets={}; self.column_config=Cfg(); self.sidebar=self
    def __enter__(self): return self
    def __exit__(self,*a): return False
    def __getattr__(self,n):
        if n in {'columns','tabs'}:
            def many(spec,*a,**k):
                count=len(spec) if isinstance(spec,(list,tuple)) else int(spec)
                return [self for _ in range(count)]
            return many
        if n=='cache_resource':
            return lambda *a,**k: (lambda fn: fn)
        if n in {'form','expander','spinner','chat_message','container'}:
            return lambda *a,**k:self
        if n=='selectbox':
            def f(label,options,*a,**k):
                opts=list(options)
                if label=='Aktif Proje' and 'Test Project' in opts: return 'Test Project'
                if label=='ANA BÖLÜM':
                    target=os.getenv('EA_TEST_PAGE','Yönetici Dashboard')
                    return next((sec for sec,pages in SECTIONS.items() if target in pages),opts[0])
                if not opts: return None
                ix=k.get('index',0)
                return opts[ix] if 0<=ix<len(opts) else opts[0]
            return f
        if n=='radio':
            def f(label,options,*a,**k):
                opts=list(options); target=os.getenv('EA_TEST_PAGE','Yönetici Dashboard')
                return target if target in opts else opts[0]
            return f
        if n=='data_editor':
            def f(data,*a,**k): return data
            return f
        if n=='file_uploader': return lambda *a,**k:None
        if n in {'button','form_submit_button','checkbox'}: return lambda *a,**k:False
        if n=='date_input':
            def f(*a,**k): return k.get('value', a[1] if len(a)>1 else __import__('datetime').date(2026,11,15))
            return f
        if n=='number_input':
            def f(*a,**k): return k.get('value', a[3] if len(a)>3 else (a[1] if len(a)>1 and isinstance(a[1],(int,float)) else 0))
            return f
        if n in {'text_input','text_area'}:
            def f(*a,**k): return k.get('value', a[1] if len(a)>1 and isinstance(a[1],str) else '')
            return f
        if n=='chat_input': return lambda *a,**k:None
        if n=='stop': return lambda *a,**k:(_ for _ in ()).throw(Stop())
        if n=='rerun': return lambda *a,**k:None
        if n=='secrets': return {}
        if n=='session_state': return self.session_state
        return lambda *a,**k:None

with tempfile.TemporaryDirectory(prefix='ea_ui_pages_') as td:
    db=Path(td)/'ui.db'; url=f'sqlite:///{db}'
    eng=create_engine(url)
    salt=os.urandom(16).hex(); pw='StrongPass!2026_UI'; ph=hashlib.pbkdf2_hmac('sha256',pw.encode(),bytes.fromhex(salt),310000).hex()
    with eng.begin() as con:
        con.execute(text('CREATE TABLE ea_users(username TEXT PRIMARY KEY,display_name TEXT NOT NULL,role TEXT NOT NULL,salt TEXT NOT NULL,password_hash TEXT NOT NULL,active INTEGER NOT NULL DEFAULT 1,created_at TEXT NOT NULL)'))
        con.execute(text("INSERT INTO ea_users(username,display_name,role,salt,password_hash,active,created_at) VALUES('admin','Admin','Yönetici',:s,:h,1,'2026-11-15')"),{'s':salt,'h':ph})
        con.execute(text('CREATE TABLE ea_records(project_key TEXT NOT NULL,dataset TEXT NOT NULL,record_id TEXT NOT NULL,payload TEXT NOT NULL,updated_at TEXT NOT NULL,PRIMARY KEY(project_key,dataset,record_id))'))
        rec={'_id':'project-1','Proje Adı':'Test Project','İşveren':'Test','Lokasyon':'Adana','Para Birimi':'TRY','Başlangıç Tarihi':'2026-01-01','Bitiş Tarihi':'2027-01-01'}
        con.execute(text("INSERT INTO ea_records(project_key,dataset,record_id,payload,updated_at) VALUES('__GLOBAL__','project','project-1',:p,'2026-11-15')"),{'p':json.dumps(rec,ensure_ascii=False)})
    eng.dispose()
    failures=[]
    pages=[p for ps in SECTIONS.values() for p in ps]
    for page in pages:
        fake=UI(); fake.session_state['auth_user']={'username':'admin','display_name':'Admin','role':'Yönetici'}
        sys.modules['streamlit']=fake
        os.environ['DATABASE_URL']=url
        os.environ['EA_TEST_PAGE']=page
        spec=importlib.util.spec_from_file_location('app_ui_smoke_'+str(pages.index(page)),APP_DIR/'app.py')
        mod=importlib.util.module_from_spec(spec); sys.modules[spec.name]=mod
        try:
            spec.loader.exec_module(mod)
            print('PASS',page)
        except Stop:
            print('STOP',page,'(st.stop expected/guard)')
        except Exception as e:
            failures.append((page,type(e).__name__,str(e)))
            print('FAIL',page,type(e).__name__,str(e))
        finally:
            try: mod.ENGINE.dispose()
            except Exception: pass
    print('\nTOTAL',len(pages),'FAILURES',len(failures))
    for f in failures: print('FAIL_DETAIL',repr(f))
