from __future__ import annotations
import os, sys, types, tempfile, importlib.util, hashlib, json, shutil, io, math, traceback
from pathlib import Path
from datetime import date, datetime
import pandas as pd
from sqlalchemy import create_engine, text
from openpyxl import Workbook
import ezdxf

SRC=Path(__file__).resolve().parents[1]/'app.py'
REPORT=[]

def record(name, ok, detail=''):
    status = 'SKIP' if ok is None else ('PASS' if ok else 'FAIL')
    REPORT.append((name, status, str(detail)))
    print(status, name, detail)

class Stop(BaseException): pass
class State(dict):
    def __getattr__(self,k):
        try: return self[k]
        except KeyError: raise AttributeError(k)
    def __setattr__(self,k,v): self[k]=v
class Cfg:
    def __getattr__(self,n): return lambda *a,**k: object()
class UI:
    def __init__(self):
        self.session_state=State(); self.secrets={}; self.column_config=Cfg(); self.sidebar=self
        self.page='Yönetici Dashboard'
    def __enter__(self): return self
    def __exit__(self,*a): return False
    def cache_resource(self,*args,**kwargs):
        if args and callable(args[0]): return args[0]
        return lambda fn: fn
    def __getattr__(self,n):
        if n in {'columns','tabs'}:
            def many(spec,*a,**k):
                count=len(spec) if isinstance(spec,(list,tuple)) else int(spec)
                return [self for _ in range(count)]
            return many
        if n in {'form','expander','spinner','chat_message','container'}: return lambda *a,**k:self
        if n=='selectbox':
            def f(label,options,*a,**k):
                opts=list(options)
                if not opts: return None
                if label=='Aktif Proje': return 'Adana Villa E2E' if 'Adana Villa E2E' in opts else opts[0]
                if label=='ANA BÖLÜM':
                    for sec,pages in self.sections.items():
                        if self.page in pages: return sec
                ix=k.get('index',0)
                return opts[ix] if isinstance(ix,int) and 0<=ix<len(opts) else opts[0]
            return f
        if n=='radio': return lambda label,options,*a,**k: self.page if self.page in list(options) else list(options)[0]
        if n=='data_editor': return lambda data,*a,**k:data
        if n=='file_uploader': return lambda *a,**k:None
        if n in {'button','form_submit_button','checkbox'}: return lambda *a,**k:False
        if n=='date_input': return lambda *a,**k:k.get('value', a[1] if len(a)>1 else date(2026,11,15))
        if n=='number_input':
            def f(*a,**k):
                if 'value' in k: return k['value']
                return a[3] if len(a)>3 else (a[1] if len(a)>1 and isinstance(a[1],(int,float)) else 0)
            return f
        if n in {'text_input','text_area'}: return lambda *a,**k:k.get('value',a[1] if len(a)>1 and isinstance(a[1],str) else '')
        if n=='chat_input': return lambda *a,**k:None
        if n=='stop': return lambda *a,**k:(_ for _ in ()).throw(Stop())
        if n=='rerun': return lambda *a,**k:None
        if n=='secrets': return {}
        if n=='session_state': return self.session_state
        return lambda *a,**k:None

class Upload:
    def __init__(self,name,data): self.name=name; self._data=data; self.size=len(data)
    def getvalue(self): return self._data
    def read(self,*a): return self._data
    def seek(self,*a): return None

# Isolated working copy ensures test data cannot affect release package or a live DB.
work=Path(tempfile.mkdtemp(prefix='ea_project_e2e_'))
app_dir=work/'app'; app_dir.mkdir(); shutil.copy2(SRC,app_dir/'app.py')
# Copy all first-party runtime modules used by the multi-provider AI and control monitor.
for helper in ('ai_gateway.py','operations_monitor.py','security_config.py'):
    helper_src=SRC.parent/helper
    if not helper_src.is_file():
        raise FileNotFoundError(f'Release helper module missing: {helper}')
    shutil.copy2(helper_src,app_dir/helper)
if str(app_dir) not in sys.path: sys.path.insert(0,str(app_dir))
db=work/'e2e.db'
os.environ['DATABASE_URL']=f'sqlite:///{db.as_posix()}'
os.environ.pop('EA_BOOTSTRAP_TOKEN',None)
os.environ.pop('OPENAI_API_KEY',None)
os.environ['EA_MAX_UPLOAD_MB']='100'
ui=UI(); ui.sections={
 'Yönetim Paneli':['Yönetici Dashboard','Şirket Portföyü','AI Proje Asistanı','Proje Performansı','Operasyon Kontrol Merkezi'],
 'Proje Yönetimi':['Proje Künyesi','Doküman Merkezi','AI Belge Analizi','Günlük Şantiye Raporu'],
 'Metraj ve Bütçe':['Excel Metraj Import','DXF / DWG Analizi','Birim Fiyat Kütüphanesi','Bütçe ve Metraj','Metraj Onayları','Veri Kontrolü'],
 'Maliyet ve Finans':['Gerçekleşen Maliyet','Fatura Analizi','EVM Performans','Nakit Akışı / S-Curve','Proje Kârlılığı'],
 'Planlama ve Saha':['İş Programı / CPM','İşçilik ve Puantaj','Risk Yönetimi'],
 'Satın Alma ve Depo':['Satın Alma ve Tedarik','Stok Hareketleri'],
 'Sözleşme ve Hakediş':['Sözleşme ve Değişiklik','Taşeron Hakediş','İşveren Hakedişi','Onay Merkezi'],
 'Sistem Yönetimi':['Kullanıcı Yönetimi','Yedekleme / Denetim','Excel Raporları','Sistem Bilgileri']}
sys.modules['streamlit']=ui
# First run creates SQLite schema then stops at secure bootstrap. Create test users via app's real auth functions.
spec=importlib.util.spec_from_file_location('ea_project_e2e_app',app_dir/'app.py')
app=importlib.util.module_from_spec(spec); sys.modules[spec.name]=app
try: spec.loader.exec_module(app)
except Stop: pass
app.create_user('admin01','Test Yöneticisi','Yönetici','StrongPass!2026_Admin')
app.create_user('pm001','Test Proje Müdürü','Proje Müdürü','StrongPass!2026_Manager')
admin=app.authenticate('admin01','StrongPass!2026_Admin')
record('Test DB ilk kurulum ve yönetici kimlik doğrulama',admin is not None)
record('ISO ve Türkçe tarih biçimlerinin uyarısız normalizasyonu', app.parse_date_value('2026-11-15')==date(2026,11,15) and app.parse_date_value('15.11.2026')==date(2026,11,15) and app.parse_date_value('2026-11-15T14:43:07+03:00')==date(2026,11,15))
ui.session_state['auth_user']=admin

# project card in the global project dataset
project_rec={'_id':'project-e2e-001','Proje Adı':'Adana Villa E2E','İşveren':'E2E Test İşvereni','Lokasyon':'Adana, Türkiye','Sözleşme No':'EA-E2E-2026','Sözleşme Tarihi':date(2026,10,1),'Başlangıç Tarihi':date(2026,11,1),'Bitiş Tarihi':date(2027,6,30),'Para Birimi':'TRY','KDV Oranı':0.20,'Risk Rezervi (%)':0.05,'Proje Yöneticisi':'Test Proje Müdürü','Açıklama':'Sentetik kabul test projesi; gerçek müşteri verisi değildir.'}
app.save_ds('__GLOBAL__','project',pd.DataFrame([project_rec]),app.PROJECT)
app.assign_project('Adana Villa E2E','pm001','admin01')
record('Proje kartı oluşturma ve proje erişim ataması',app.can_project('pm001','Proje Müdürü','Adana Villa E2E'))
# Re-execute whole Streamlit script with seeded admin session; this checks actual startup + main dashboard path.
ui.session_state['current_project']='Adana Villa E2E'; ui.page='Yönetici Dashboard'
spec2=importlib.util.spec_from_file_location('ea_project_e2e_ui',app_dir/'app.py')
ui_mod=importlib.util.module_from_spec(spec2); sys.modules[spec2.name]=ui_mod
try:
    spec2.loader.exec_module(ui_mod)
    app=ui_mod
    record('Gerçek uygulama başlangıç akışı / yönetici paneli kabuğu',True,'headless Streamlit stub; gerçek tarayıcı değil')
except Exception as e:
    record('Gerçek uygulama başlangıç akışı / yönetici paneli kabuğu',False,repr(e)); raise

# Create a realistic synthetic project workbook.
wb=Workbook(); ws=wb.active; ws.title='Metraj'
headers=['WBS','Poz No','İş Kalemi','Ana Kategori','Birim','Miktar','Birim Fiyat','Para Birimi','Kur','Fiyat Yılı','Fiyat Kaynağı','Metraj Durumu']
ws.append(headers)
items=[
 ['KAB.01','EA-BET-001','Grobeton C16/20','Betonarme İmalatlar','m3',25,2400,'TRY',1,2026,'Test teklif / 2026','Kontrol Bekliyor'],
 ['KAB.02','EA-BET-030','C30/37 temel betonu','Betonarme İmalatlar','m3',180,4200,'TRY',1,2026,'Test teklif / 2026','Kontrol Bekliyor'],
 ['KAB.03','EA-DEM-014','B420C nervürlü inşaat demiri Ø14 mm','Betonarme İmalatlar','ton',22,34000,'TRY',1,2026,'Test teklif / 2026','Kontrol Bekliyor'],
 ['YAL.01','EA-YAL-001','Temel altı bitümlü membran 4 mm','Yalıtım / Çatı / Cephe','m2',250,550,'TRY',1,2026,'Test teklif / 2026','Kontrol Bekliyor'],
 ['INC.01','EA-SER-060','60x60 porselen seramik zemin kaplama','Mimari İnce İşler','m2',120,950,'TRY',1,2026,'Test teklif / 2026','Kontrol Bekliyor'],
 ['CEP.01','EA-CEP-001','Alüminyum kompozit cephe kaplama','Yalıtım / Çatı / Cephe','m2',180,3500,'TRY',1,2026,'Test teklif / 2026','Kontrol Bekliyor'],
 ['MEK.01','EA-MEK-025','PPRC temiz su borusu Ø25 mm','Mekanik ve Elektrik','m',140,190,'TRY',1,2026,'Test teklif / 2026','Kontrol Bekliyor'],
 ['ELE.01','EA-ELE-325','NYM elektrik kablosu 3x2.5 mm','Mekanik ve Elektrik','m',500,85,'TRY',1,2026,'Test teklif / 2026','Kontrol Bekliyor'],
]
for r in items: ws.append(r)
ws2=wb.create_sheet('Is_Programi'); ws2.append(['WBS','Aktivite','Süre (gün)','Öncül WBS','Bağımlılık Tipi','Başlangıç','Bitiş','Tamamlanma (%)','Ağırlık (%)','Sorumlu','Durum'])
for i,r in enumerate(items):
    w=r[0]; pred=items[i-1][0] if i else ''
    ws2.append([w,r[2],30,pred,'FS' if pred else '',date(2026,11,1),date(2026,12,31),25 if i<3 else 0,1,'Test Ekip','Devam Ediyor'])
ws3=wb.create_sheet('Fatura_Demir'); ws3.append(['WBS','Poz No','İş Kalemi','Birim','Gerçekleşen Miktar','Gerçekleşen Birim Fiyat','Para Birimi','Kur','Tarih','Tedarikçi','Belge No','Kayıt Türü','Durum'])
ws3.append(['KAB.03','EA-DEM-014','B420C nervürlü inşaat demiri Ø14 mm','kg',5000,34.5,'TRY',1,date(2026,11,10),'Adana Demir Tedarik','FAT-E2E-001','Fatura','Kontrol Bekliyor'])
workbook=work/'Adana_Villa_Test_Projesi.xlsx'; wb.save(workbook)
blob=workbook.read_bytes(); upload=Upload(workbook.name,blob)
name,digest=app.save_uploaded('Adana Villa E2E',upload,'admin01')
docs=app.list_documents('Adana Villa E2E')
record('Örnek proje Excel dosyasını güvenli doküman deposuna yükleme',len(docs)==1 and docs.iloc[0]['original_name']==workbook.name, f'{len(docs)} belge kaydı')
try:
    app.save_uploaded('Adana Villa E2E',upload,'admin01')
    dup_ok=False
except ValueError as e:
    dup_ok='daha önce yüklenmiş' in str(e)
record('Aynı proje dosyasının mükerrer yüklenmesini engelleme',dup_ok)
# Import workbook sheets using the application's real mapping and normalization functions.
sheets=pd.read_excel(io.BytesIO(blob),sheet_name=None)
parsed_invoice=app.map_excel(sheets['Fatura_Demir'],'actual')
record('Excel fatura satırının gerçekleşen adayı olarak okunması',len(parsed_invoice)==1 and parsed_invoice.iloc[0]['WBS']=='KAB.03' and math.isclose(parsed_invoice.iloc[0]['Gerçekleşen Miktar'],5000))
plan_raw=sheets['Metraj']; takeoff_mapping={'WBS':'WBS','İş Kalemi':'İş Kalemi','Birim':'Birim','Miktar':'Miktar','Poz No':'Poz No','Para Birimi':'Para Birimi','Birim Fiyat':'Birim Fiyat','Fiyat Yılı':'Fiyat Yılı','Fiyat Kaynağı':'Fiyat Kaynağı'}
takeoff_preview=app.build_excel_takeoff_rows(plan_raw,takeoff_mapping,'TR (1.234,56)','TRY',pd.DataFrame(columns=app.LIB),'Adana_Villa_Test_Projesi.xlsx','Metraj')
record('Excel ekranındaki ortak metraj hazırlama motoru ve WBS alanı',len(takeoff_preview)==8 and set(takeoff_preview['WBS'])=={r[0] for r in items})
plan_import=app.assign_import_wbs(takeoff_preview,pd.DataFrame(columns=app.PLAN))
record('Excel metraj aktarımında WBS kodlarının korunması',set(plan_import['WBS'])=={r[0] for r in items})
lowercase_rows=takeoff_preview.copy(); lowercase_rows.loc[lowercase_rows.index[0],'WBS']='kab.01'
normalized_rows=app.assign_import_wbs(lowercase_rows,pd.DataFrame(columns=app.PLAN))
record('WBS kodlarının büyük harfe normalize edilmesi',normalized_rows.iloc[0]['WBS']=='KAB.01')
record('Excel birim fiyatı, yılı ve kaynak bilgisinin korunması',plan_import['Birim Fiyat'].notna().all() and set(plan_import['Fiyat Yılı'].astype(int))=={2026} and plan_import['Fiyat Kaynağı'].astype(str).str.contains('Test teklif').all())
try:
    app.assign_import_wbs(pd.concat([takeoff_preview.iloc[[0]],takeoff_preview.iloc[[0]]],ignore_index=True),pd.DataFrame(columns=app.PLAN)); dup_wbs_blocked=False
except ValueError as e: dup_wbs_blocked='tekrarlı WBS' in str(e)
record('Excel içinde yinelenen WBS kodlarının engellenmesi',dup_wbs_blocked)
try:
    app.assign_import_wbs(takeoff_preview,plan_import[app.PLAN]); existing_wbs_blocked=False
except ValueError as e: existing_wbs_blocked='zaten mevcut' in str(e)
record('Projede mevcut WBS kodlarının mükerrer importunun engellenmesi',existing_wbs_blocked)
plan=app.ensure_frame(plan_import[app.PLAN],app.PLAN)
plan['_id']=[app.uid() for _ in range(len(plan))]
plan['Metraj Durumu']='Kontrol Bekliyor'
plan=app.calc_plan(plan,'TRY')
schedule=app.map_excel(sheets['Is_Programi'],'schedule'); schedule['_id']=[app.uid() for _ in range(len(schedule))]
app.save_ds('Adana Villa E2E','plan',plan,app.PLAN)
app.save_ds('Adana Villa E2E','schedule',schedule,app.SCHEDULE)
loaded_plan=app.load_ds('Adana Villa E2E','plan',app.PLAN)
record('Excel metraj sayfasının WBS ve imalat sınıflarıyla okunması',len(loaded_plan)==8 and loaded_plan['WBS'].nunique()==8)
record('İş programı sayfasının bütçe WBS kodlarıyla okunması',len(schedule)==8 and set(schedule.WBS)==set(loaded_plan.WBS))
record('Bütçe tutarlarının hesaplanması',math.isclose(float(plan['Planlanan Tutar'].sum()),float(sum(r[5]*r[6] for r in items)),abs_tol=.01),f"Toplam {plan['Planlanan Tutar'].sum():,.2f} TRY")
# Classification must resolve the intended groups.
classifications={r['WBS']:app.classify_work_item(r['İş Kalemi'],r['Ana Kategori']) for _,r in plan.iterrows()}
record('Kaba / ince / cephe / mekanik / elektrik / yalıtım sınıflandırması',
       classifications['KAB.01'][0]=='Kaba İşler' and classifications['KAB.03'][1]=='Donatı / İnşaat Demiri' and classifications['YAL.01'][1]=='Temel / Perde Su Yalıtımı' and classifications['INC.01'][0]=='İnce İşler' and classifications['CEP.01'][0]=='Cephe İşleri' and classifications['MEK.01'][0]=='Mekanik Tesisat' and classifications['ELE.01'][0]=='Elektrik Tesisat',repr(classifications))
# Price library test for steel and cement.
lib=pd.DataFrame([{ '_id':'lib-steel','Poz No':'EA-DEM-014','İş Kalemi':'B420C nervürlü inşaat demiri Ø14 mm','Birim':'ton','Birim Fiyat':34000,'Para Birimi':'TRY','Fiyat Yılı':2026,'Kaynak':'Test tedarikçi teklifi','Aktif':True,'Geçerlilik Başlangıcı':date(2026,1,1),'Yayın Tarihi':date(2026,1,1)}])
app.save_ds('__GLOBAL__','library',lib,app.LIB)
steel_row=plan[plan.WBS=='KAB.03'].iloc[0].to_dict(); hit=app.price_match(steel_row,lib)
record('Gerçek fiyat kütüphanesi eşleştirmesi',hit.get('status')=='YÜKSEK EŞLEŞME — ONAY GEREKLİ' and hit.get('price')==34000,hit.get('status'))
# Approval flow for budget lines: PM requests, manager approves.
ui.session_state['auth_user']=app.authenticate('pm001','StrongPass!2026_Manager')
approval_ids=[]
for _,r in plan.iterrows():
    approval_ids.append(app.submit_approval('Adana Villa E2E','plan',r['_id'],r['İş Kalemi'],r['Planlanan Tutar'],'TRY','Yönetici'))
record('Bütçe satırlarının çoklu onay talebine gönderilmesi',len(approval_ids)==8)
ui.session_state['auth_user']=app.authenticate('admin01','StrongPass!2026_Admin')
for aid in approval_ids: app.decide_approval('Adana Villa E2E',aid,'Onaylandı','E2E kabul testi')
approved_plan=app.load_ds('Adana Villa E2E','plan',app.PLAN)
record('Yönetici tarafından bütçe onayı ve denetim izi',approved_plan['Metraj Durumu'].map(app.is_approved).all() and len(app.load_ds('Adana Villa E2E','approvals',app.APPROVAL))==8)
# Create a valid immutable approved baseline.
baselines,baseline_row=app.build_budget_schedule_baseline(approved_plan,schedule,'TRY','admin01',app.load_ds('Adana Villa E2E','baseline',app.BASELINE))
app.save_ds('Adana Villa E2E','baseline',baselines,app.BASELINE)
verified_s,verified_p,meta=app.verify_active_baseline(baseline_row,'TRY')
record('Onaylı bütçe + iş programı baseline oluşturma ve hash doğrulama',len(verified_p)==8 and math.isclose(meta['BudgetTotal'],float(approved_plan['Planlanan Tutar'].sum()),abs_tol=.01))
# Purchase invoice to actual; 5000kg @34.50 should post as 5ton @34500 without changing total.
purchase_row={'_id':'po-e2e-1','Sipariş No':'PO-E2E-001','WBS':'KAB.03','Poz No':'EA-DEM-014','Malzeme / Hizmet':'B420C nervürlü inşaat demiri Ø14 mm','Tedarikçi':'Adana Demir Tedarik','Birim':'kg','Faturalanan Miktar':5000,'Faturalanan Tutar':172500,'Para Birimi':'TRY','Kur':1,'Kur Kaynağı':'Proje baz para birimi','Fatura No':'FAT-E2E-001','Fatura Tarihi':date(2026,11,10)}
actual_row=app.purchase_invoice_to_actual(purchase_row,approved_plan,'TRY')
record('Demir faturası: 5.000 kg → 5 ton ve toplam tutar korunması',actual_row['Birim']=='ton' and math.isclose(actual_row['Gerçekleşen Miktar'],5) and math.isclose(actual_row['Gerçekleşen Birim Fiyat'],34500) and math.isclose(actual_row['Gerçekleşen Tutar'],172500),str({k:actual_row[k] for k in ['WBS','Birim','Gerçekleşen Miktar','Gerçekleşen Birim Fiyat','Gerçekleşen Tutar']}))
actuals=app.calc_actual(pd.DataFrame([actual_row]),'TRY')
app.save_ds('Adana Villa E2E','actual',actuals,app.ACTUAL)
# Pending AC must be zero until approved.
evm_pending=app.calculate_evm(approved_plan,actuals,schedule,date(2026,11,15),'TRY',verified_s,pd.DataFrame(),verified_p)
record('Gerçekleşen onaylanmadan EVM AC dışı kalması',(float(evm_pending['AC'])==0 and int(evm_pending['Onaylı gerçekleşen satırı'])==0))
# actual approval
ui.session_state['auth_user']=app.authenticate('pm001','StrongPass!2026_Manager')
aid=app.submit_approval('Adana Villa E2E','actual',actual_row['_id'],'Demir alımı',172500,'TRY','Yönetici')
ui.session_state['auth_user']=app.authenticate('admin01','StrongPass!2026_Admin')
app.decide_approval('Adana Villa E2E',aid,'Onaylandı','Fatura ve miktar doğrulandı')
actuals2=app.load_ds('Adana Villa E2E','actual',app.ACTUAL)
evm=app.calculate_evm(approved_plan,actuals2,schedule,date(2026,11,15),'TRY',verified_s,pd.DataFrame(),verified_p)
record('Gerçekleşen onayından sonra AC hesabına dahil edilmesi',math.isclose(float(evm['AC']),172500,abs_tol=.01) and int(evm['Onaylı gerçekleşen satırı'])==1, f"AC={evm['AC']}")
# WBS quantity and cost comparison
variance=app.work_item_variance(approved_plan,actuals2,'TRY')
steel_var=variance[variance.WBS=='KAB.03'].iloc[0]
record('WBS bazında bütçe/gerçekleşen demir karşılaştırması',math.isclose(steel_var['Gerçekleşen Miktar'],5) and math.isclose(steel_var['Gerçekleşen Tutar'],172500) and steel_var['Teknik Kontrol']=='UYGUN',str(steel_var[['Bütçe Miktarı','Gerçekleşen Miktar','Bütçe Tutarı','Gerçekleşen Tutar','Teknik Kontrol']].to_dict()))
# Actual summary/dashboard metrics and control findings
summary=app.summarize_costs(approved_plan,actuals2,'TRY'); findings=app.project_control_findings(approved_plan,actuals2,schedule,pd.DataFrame(columns=app.RISK),date(2026,11,15))
record('Ana imalat grubu yönetim özeti üretilmesi',not app.work_group_summary(approved_plan,actuals2,'TRY').empty)
record('Otomatik proje kontrol bulguları üretilmesi',not findings.empty)
# Cash FX conversion and employer progress payment math
cash=pd.DataFrame([{'_id':'cash1','Tarih':date(2026,11,11),'Tür':'Ödeme','Açıklama':'Demir ödemesi','Firma / Cari':'Adana Demir','Tutar':172500,'Para Birimi':'TRY','Kur':1,'Kur Kaynağı':'TRY','Baz Para Birimi Tutarı':172500,'Vade Tarihi':date(2026,11,20),'Durum':'Gerçekleşti','WBS':'KAB.03','Belge No':'FAT-E2E-001'}])
app.save_ds('Adana Villa E2E','cash',cash,app.CASH)
record('Nakit akışı kaydı saklama ve yükleme',len(app.load_ds('Adana Villa E2E','cash',app.CASH))==1)
hakedis=pd.DataFrame([{'_id':'hak1','Brüt İmalat Tutarı':100000,'Avans Mahsubu':10000,'Teminat Kesintisi':5000,'Diğer Kesintiler':0,'KDV Oranı':20,'Tahsil Edilen':50000}])
hakedis_calc=app.employer_hakedis_calc(hakedis).iloc[0]
record('İşveren hakediş kesinti/KDV/tahsilat hesabı',math.isclose(hakedis_calc['Hakediş Toplamı'],102000) and math.isclose(hakedis_calc['Kalan Alacak'],52000),str(hakedis_calc[['Net Tahakkuk','KDV Tutarı','Hakediş Toplamı','Kalan Alacak']].to_dict()))
# Create/upload real DXF fixture (not a true building drawing, but geometry fixture).
doc=ezdxf.new('R2010'); doc.header['$INSUNITS']=4; msp=doc.modelspace()
doc.layers.new('BETON_GROBETON'); msp.add_lwpolyline([(0,0),(10000,0),(10000,8000),(0,8000)],close=True,dxfattribs={'layer':'BETON_GROBETON'})
doc.layers.new('SERAMIK_ZEMIN'); msp.add_lwpolyline([(20000,0),(26000,0),(26000,5000),(20000,5000)],close=True,dxfattribs={'layer':'SERAMIK_ZEMIN'})
doc.layers.new('DONATI_TEST'); msp.add_line((0,0),(12000,0),dxfattribs={'layer':'DONATI_TEST'})
dxf_path=work/'Villa_Test_Pafta.dxf'; doc.saveas(dxf_path)
dxf_upload=Upload(dxf_path.name,dxf_path.read_bytes())
app.save_uploaded('Adana Villa E2E',dxf_upload,'admin01')
dxf_result=app.read_dxf(dxf_path)
record('DXF dosya yükleme ve geometri okuma',len(dxf_result['detail'])==3 and dxf_result['unit_code']==4)
areas=dxf_result['layers'].get('BETON_GROBETON',{}).get('Alan',0)
record('DXF ölçekli alan hesabı (mm → m²)',math.isclose(areas*0.001*0.001,80,rel_tol=.02),f"test alanı yaklaşık {areas*1e-6:.2f} m²".replace('.',','))
dxf_candidates=app.dxf_layer_takeoff_candidates(dxf_result,lib,'TRY','Villa_Test_Pafta.dxf','sha256-e2e-test')
seramik_candidate=dxf_candidates[dxf_candidates['Kaynak Layer'].astype(str)=='SERAMIK_ZEMIN'] if not dxf_candidates.empty else pd.DataFrame()
record('DXF katmanından seramik alan metraj adayı üretimi',len(seramik_candidate)==1 and seramik_candidate.iloc[0]['Birim']=='m²' and math.isclose(float(seramik_candidate.iloc[0]['Miktar']),30,rel_tol=.02),str(seramik_candidate[['Birim','Miktar']].to_dict('records') if not seramik_candidate.empty else []))
record('DXF aynı dosya/katman tekrar aktarım kontrolü',app.find_duplicate_dxf_layers(approved_plan,dxf_candidates,'sha256-e2e-test')==[] and app.find_duplicate_dxf_layers(pd.concat([approved_plan,pd.DataFrame([{'Açıklama':'DXF katmanı: SERAMIK_ZEMIN · DXF SHA256: sha256-e2e-test'}])],ignore_index=True),dxf_candidates,'sha256-e2e-test')==['SERAMIK_ZEMIN'])
# Verify unsupported integrations fail safely instead of pretending to work.
try: app.convert_dwg(work/'not-real.dwg',work/'out'); dwg_safe=False
except RuntimeError as e: dwg_safe='varsayılan olarak kapalı' in str(e)
record('DWG dönüştürücü yapılandırılmadığında güvenli durması',dwg_safe)
if importlib.util.find_spec('ifcopenshell') is None:
    try:
        app.read_ifc(work/'dummy.ifc')
        record('IFC opsiyonel bağımlılık güvenli hata kontrolü',False,'ifcopenshell yokken hata vermesi beklenirdi')
    except RuntimeError as e:
        record('IFC opsiyonel bağımlılık güvenli hata kontrolü','opsiyonel ifcopenshell' in str(e),str(e))
    record('Gerçek IFC dosyası uçtan uca entegrasyon testi',None,'ifcopenshell bu test ortamında kurulu değil; gerçek IFC dosyası doğrulanamadı')
else:
    record('IFC motoru kullanılabilirliği',True,'ifcopenshell mevcut')
try:
    app.ai_client()
    record('AI anahtar/model yapılandırma kontrolü',True,'ai_client yapılandırılmış')
except RuntimeError as e:
    record('AI sağlayıcı anahtar/model yapılandırma kontrolü','Yapay zekâ sağlayıcısı yapılandırılmamış' in str(e),str(e))
    record('Canlı AI API uçtan uca testi',None,'OpenAI/Claude/Gemini anahtarları bu test ortamında tanımlı değil; canlı AI çağrısı yapılmadı')
# Render every page against populated project datasets through headless UI shell.
page_failures=[]
all_pages=[p for page_list in ui.sections.values() for p in page_list]
for ix,page_name in enumerate(all_pages):
    ui.page=page_name
    ui.session_state['auth_user']=app.authenticate('admin01','StrongPass!2026_Admin')
    ui.session_state['current_project']='Adana Villa E2E'
    pg_spec=importlib.util.spec_from_file_location(f'ea_project_e2e_page_{ix}',app_dir/'app.py')
    pg_mod=importlib.util.module_from_spec(pg_spec); sys.modules[pg_spec.name]=pg_mod
    try:
        pg_spec.loader.exec_module(pg_mod)
    except Exception as e:
        page_failures.append((page_name,type(e).__name__,str(e)))
    finally:
        try: pg_mod.ENGINE.dispose()
        except Exception: pass
record(f'Dolu test projesiyle tüm {len(all_pages)} sayfa yönlendirme/render duman testi',len(page_failures)==0,(f'{len(all_pages)} sayfa' if not page_failures else repr(page_failures[:5])))
# Rebind the final module after page smoke tests to a fresh app import for report export.
ui.page='Yönetici Dashboard'; ui.session_state['auth_user']=app.authenticate('admin01','StrongPass!2026_Admin')
last_spec=importlib.util.spec_from_file_location('ea_project_e2e_final',app_dir/'app.py')
last_mod=importlib.util.module_from_spec(last_spec); sys.modules[last_spec.name]=last_mod; last_spec.loader.exec_module(last_mod); app=last_mod
# Export report workbook and check readable output
xlsx_bytes=app.excel_export({'Bütçe':approved_plan,'Gerçekleşen':actuals2,'Mutabakat':variance,'Kontrol':findings})
out_report=work/'Adana_Villa_E2E_Kabul_Raporu.xlsx'; out_report.write_bytes(xlsx_bytes)
export_sheets=pd.read_excel(out_report,sheet_name=None)
record('Excel yönetim raporu çıktısı',set(export_sheets)=={'Bütçe','Gerçekleşen','Mutabakat','Kontrol'})
# Project archive with two input files
project_zip=work/'Adana_Villa_E2E_Test_Projesi.zip'
import zipfile
with zipfile.ZipFile(project_zip,'w',compression=zipfile.ZIP_DEFLATED) as z:
    z.write(workbook,workbook.name); z.write(dxf_path,dxf_path.name)

# Tests summary and outputs
passed=sum(1 for _,status,_ in REPORT if status=='PASS'); failed=sum(1 for _,status,_ in REPORT if status=='FAIL'); skipped=sum(1 for _,status,_ in REPORT if status=='SKIP')
lines=['# Adana Villa E2E — Proje Kabul Testi','',f'- Tarih: {datetime.now().strftime("%d.%m.%Y %H:%M")}', '- Test verisi: Sentetik örnek proje; gerçek müşteri projesi değildir.', f'- Sonuç: **{passed} geçti, {failed} başarısız, {skipped} entegrasyon testi çalıştırılamadı**','', '## Test detayları','']
for n,status,d in REPORT: lines.append(f'- [{status}] {n}'+(f' — {d}' if d else ''))
lines += ['', '## Sınırlar','', '- Testler gerçek uygulama Python iş fonksiyonları, izole SQLite veritabanı ve taklit Streamlit kabuğu üzerinde çalıştırıldı. Gerçek tarayıcı ve Streamlit Cloud testi değildir.', '- DXF dosyası geometrik test fikstürüdür; gerçek mimari/statik pafta değildir.', '- DWG dönüştürmesi ODA File Converter bulunmadığı için çalıştırılmadı; güvenli varsayılan kapalılık test edildi.', '- ifcopenshell bu ortamda kurulu olmadığından gerçek IFC dosyası işlenemedi.', '- OpenAI/Anthropic/Gemini API anahtarları bulunmadığından canlı AI API isteği yapılmadı.', '- YFK canlı HTTP/PDF fiyat senkronizasyonu internet erişimi ve güncel kaynağa bağlı olduğundan bu kabul testinde çalıştırılmadı.', '- KDV dahil/hariç alan ayrımı ürün kapsamı açığı olarak kaldı.']
report=work/'E2E_PROJECT_ACCEPTANCE_TR.md'; report.write_text('\n'.join(lines),encoding='utf-8')
print('\nSUMMARY',passed,'passed;',failed,'failed;',skipped,'skipped')
print('WORKDIR',work)
print('PROJECT_ZIP',project_zip)
print('REPORT',report)
if failed:
    for n,status,d in REPORT:
        if status=='FAIL': print('FAILED:',n,d)
    raise SystemExit(1)
