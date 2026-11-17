"""Runtime smoke tests for startup, SQLite persistence, authentication and write-conflict protection.

This intentionally stubs only Streamlit's UI shell so the tests can run in a
headless CI/container. It does NOT claim to test browser rendering or widgets.
"""
from __future__ import annotations

import importlib.util
import json
import os
import sys
import tempfile
import zipfile
import types
import unittest
from pathlib import Path

import pandas as pd


class _BootstrapStop(BaseException):
    """Represents st.stop() during the expected first-run bootstrap screen."""


class _SessionState(dict):
    pass


def _load_app_with_streamlit_stub(app_path: Path, database_path: Path):
    fake_st = types.ModuleType("streamlit")
    fake_st.session_state = _SessionState()
    fake_st.secrets = {}

    def cache_resource(*args, **kwargs):
        def decorate(fn):
            return fn
        if args and callable(args[0]):
            return args[0]
        return decorate

    def stop(*args, **kwargs):
        raise _BootstrapStop()

    def ui_noop(*args, **kwargs):
        return None

    fake_st.cache_resource = cache_resource
    fake_st.stop = stop
    for name in (
        "set_page_config", "markdown", "title", "error", "warning", "info",
        "success", "caption", "write", "subheader", "divider", "header",
        "metric", "toast", "rerun", "experimental_rerun",
    ):
        setattr(fake_st, name, ui_noop)

    sys.modules["streamlit"] = fake_st
    os.environ["DATABASE_URL"] = f"sqlite:///{database_path.as_posix()}"
    os.environ.pop("EA_BOOTSTRAP_TOKEN", None)

    # The release is a multi-file package; resolve sibling modules exactly as Streamlit does.
    app_dir = str(app_path.resolve().parent)
    if app_dir not in sys.path:
        sys.path.insert(0, app_dir)
    spec = importlib.util.spec_from_file_location("ea_v15_runtime_smoke", app_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("Uygulama modülü yüklenemedi.")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
    except _BootstrapStop:
        # Expected: no users exist yet, so the app shows the first-run gate.
        pass
    return module, fake_st


class RuntimeSmokeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.TemporaryDirectory(prefix="ea_v15_runtime_")
        cls.root = Path(cls.temp_dir.name)
        app_path = Path(__file__).resolve().parents[1] / "app.py"
        cls.app, cls.st = _load_app_with_streamlit_stub(app_path, cls.root / "test.db")
        if cls.app.ENGINE is None or cls.app.DB_KIND != "sqlite":
            raise AssertionError("SQLite bağlantısı başlatılamadı.")
        cls.app.create_user("admin01", "Sistem Yöneticisi", "Yönetici", "StrongPass!2026_Test")
        cls.admin = {"username": "admin01", "display_name": "Sistem Yöneticisi", "role": "Yönetici"}

    @classmethod
    def tearDownClass(cls):
        cls.temp_dir.cleanup()

    def setUp(self):
        self.st.session_state.clear()
        self.st.session_state["auth_user"] = self.admin.copy()

    def _plan_row(self, record_id: str):
        row = {column: None for column in self.app.PLAN}
        row.update({
            "_id": record_id,
            "WBS": "KAB.01",
            "Poz No": "",
            "İş Kalemi": "Grobeton",
            "Ana Kategori": "Betonarme İmalatlar",
            "Ana İş Grubu": "Kaba İşler",
            "Alt Kategori": "Grobeton",
            "Birim": "m3",
            "Miktar": 20.0,
            "Birim Fiyat": 2500.0,
            "Para Birimi": "TRY",
            "Kur": 1.0,
            "Metraj Durumu": "Kontrol Bekliyor",
        })
        return row

    def test_ai_context_compaction_preserves_valid_json_and_limit(self):
        import json
        payload={"Modül Verileri":{"plan":[{"WBS":f"WBS-{i}","Açıklama":"a"*900} for i in range(35)]}}
        compact=self.app.compact_ai_json(payload,2500)
        parsed=json.loads(compact)
        self.assertLessEqual(len(compact),2500)
        self.assertIn("Modül Verileri",parsed)
        self.assertTrue(parsed["Modül Verileri"]["plan"])

    def test_ai_multi_provider_review_requires_evidence_and_never_auto_approves(self):
        import json
        from unittest.mock import patch
        provider_results=[{"provider":"openai","model":"model-x","status":"success","output_text":json.dumps({"agent_reviews":[
            {"expert_role":"Maliyet kontrol","findings":[
                {"severity":"YÜKSEK","module":"Bütçe","record_ref":"WBS-01","issue":"Birim fiyat kaynağı eksik.","evidence":"Fiyat Kaynağı alanı boş.","action":"Kaynak ve yılı doğrulayın.","confidence":0.91},
                {"severity":"KRİTİK","module":"Risk","record_ref":"R-02","issue":"Kanıtsız bulgu.","evidence":"","action":"Kontrol edin.","confidence":1.0}
            ]}
        ]},ensure_ascii=False),"error":""}]
        with patch.object(self.app,"multi_provider_review",return_value=provider_results):
            rows,results=self.app.ai_multi_agent_review("{}",[])
        self.assertEqual(len(rows),1)
        self.assertEqual(rows.iloc[0]["Durum"],"AI ÖNERİSİ — İNSAN/UZMAN ONAYI BEKLİYOR")
        self.assertEqual(rows.iloc[0]["AI Sağlayıcı"],"openai")
        self.assertEqual(len(results),1)

    def test_settings_read_streamlit_secrets_and_environment_override(self):
        key = "EA_TEST_SETTING_SOURCE"
        prior = os.environ.get(key)
        self.st.secrets[key] = "from-streamlit-secrets"
        try:
            os.environ.pop(key, None)
            self.assertEqual(self.app.setting(key), "from-streamlit-secrets")
            os.environ[key] = "from-environment"
            self.assertEqual(self.app.setting(key), "from-environment")
        finally:
            self.st.secrets.pop(key, None)
            if prior is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = prior

    def test_invalid_upload_limit_env_uses_safe_default(self):
        key = "EA_TEST_BAD_UPLOAD_LIMIT"
        prior = os.environ.get(key)
        warnings_before = list(self.app.CONFIG_WARNINGS)
        try:
            os.environ[key] = "not-a-number"
            value = self.app._bounded_int_env(key, 100, 1, 1000)
            self.assertEqual(value, 100)
            self.assertTrue(any(key in item for item in self.app.CONFIG_WARNINGS))
        finally:
            if prior is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = prior
            self.app.CONFIG_WARNINGS[:] = warnings_before

    def test_invalid_ooxml_ratio_uses_safe_default(self):
        key = "EA_TEST_BAD_OOXML_RATIO"
        prior = os.environ.get(key)
        warnings_before = list(self.app.CONFIG_WARNINGS)
        try:
            os.environ[key] = "nan"
            value = self.app._bounded_float_env(key, 150.0, 1.0, 10000.0)
            self.assertEqual(value, 150.0)
            self.assertTrue(any(key in item for item in self.app.CONFIG_WARNINGS))
        finally:
            if prior is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = prior
            self.app.CONFIG_WARNINGS[:] = warnings_before

    def test_password_hash_and_authentication(self):
        good = self.app.authenticate("admin01", "StrongPass!2026_Test")
        bad = self.app.authenticate("admin01", "incorrect-password")
        self.assertEqual(good["role"], "Yönetici")
        self.assertIsNone(bad)

    def test_password_reset_clears_login_lockout(self):
        username = "locked01"
        self.app.create_user(username, "Kilitli Kullanıcı", "Görüntüleyici", "InitialPass!2026")
        for _ in range(5):
            self.assertIsNone(self.app.authenticate(username, "wrong-password"))
        # The account is now locked even when the correct password is supplied.
        self.assertIsNone(self.app.authenticate(username, "InitialPass!2026"))
        self.app.reset_password(username, "NewStrongPass!2026")
        authenticated = self.app.authenticate(username, "NewStrongPass!2026")
        self.assertIsNotNone(authenticated)
        self.assertEqual(authenticated["username"], username)

    def test_sqlite_dataset_round_trip_and_audit(self):
        project = "Persistence Project"
        frame = pd.DataFrame([self._plan_row("plan-001")])
        self.app.save_ds(project, "plan", frame, self.app.PLAN)
        loaded = self.app.load_ds(project, "plan", self.app.PLAN)
        self.assertEqual(len(loaded), 1)
        self.assertEqual(loaded.iloc[0]["WBS"], "KAB.01")
        self.assertAlmostEqual(float(loaded.iloc[0]["Miktar"]), 20.0)
        audit = self.app.load_ds(
            project, "audit",
            ["_id", "Tarih", "Kullanıcı", "İşlem", "Veri Seti", "Kayıt Sayısı", "Detay"],
        )
        self.assertGreaterEqual(len(audit), 1)

    def test_company_portfolio_keeps_each_project_currency_separate(self):
        projects=pd.DataFrame([
            {"_id":"prj-try","Proje Adı":"TRY Project","Para Birimi":"TRY","İşveren":"Owner A","Lokasyon":"Adana"},
            {"_id":"prj-eur","Proje Adı":"EUR Project","Para Birimi":"EUR","İşveren":"Owner B","Lokasyon":"Berlin"},
        ])
        data={}
        for name,currency,amount in [("TRY Project","TRY",10000.0),("EUR Project","EUR",2000.0)]:
            plan={c:None for c in self.app.PLAN}
            plan.update({"_id":f"p-{name}","WBS":"A.01","İş Kalemi":"Grobeton","Birim":"m³","Miktar":1,
                         "Birim Fiyat":amount,"Para Birimi":currency,"Kur":1,"Metraj Durumu":"Onaylandı"})
            actual={c:None for c in self.app.ACTUAL}
            actual.update({"_id":f"a-{name}","WBS":"A.01","İş Kalemi":"Grobeton","Birim":"m³","Gerçekleşen Miktar":1,
                           "Gerçekleşen Birim Fiyat":amount/2,"Para Birimi":currency,"Kur":1,"Tarih":"2026-11-01","Durum":"Onaylandı"})
            data[(name,"plan")]=pd.DataFrame([plan]); data[(name,"actual")]=pd.DataFrame([actual])
            data[(name,"schedule")]=pd.DataFrame(columns=self.app.SCHEDULE)
            data[(name,"risk")]=pd.DataFrame(columns=self.app.RISK)
        def loader(project,dataset,cols):
            return data.get((project,dataset),pd.DataFrame(columns=cols))
        portfolio=self.app.build_portfolio_summary(projects,loader,self.app.date(2026,11,15))
        self.assertEqual(set(portfolio["Para Birimi"]),{"TRY","EUR"})
        self.assertEqual(portfolio.set_index("Proje").loc["TRY Project","Onaylı Bütçe"],10000.0)
        self.assertEqual(portfolio.set_index("Proje").loc["EUR Project","Onaylı Bütçe"],2000.0)
        summary=self.app.portfolio_currency_summary(portfolio)
        self.assertEqual(len(summary),2)
        self.assertEqual(set(summary["Para Birimi"]),{"TRY","EUR"})
        self.assertNotIn("Genel Toplam",summary.columns)

    def test_company_portfolio_marks_missing_fx_as_incomplete_not_zero(self):
        projects=pd.DataFrame([{"_id":"prj-eur","Proje Adı":"EUR Project","Para Birimi":"TRY"}])
        plan={c:None for c in self.app.PLAN}
        plan.update({"_id":"p1","WBS":"A.01","İş Kalemi":"Grobeton","Birim":"m³","Miktar":1,
                     "Birim Fiyat":100,"Para Birimi":"EUR","Kur":None,"Metraj Durumu":"Onaylandı"})
        actual={c:None for c in self.app.ACTUAL}
        actual.update({"_id":"a1","WBS":"A.01","İş Kalemi":"Grobeton","Birim":"m³","Gerçekleşen Miktar":1,
                       "Gerçekleşen Birim Fiyat":50,"Para Birimi":"EUR","Kur":None,"Tarih":"2026-11-01","Durum":"Onaylandı"})
        data={"plan":pd.DataFrame([plan]),"actual":pd.DataFrame([actual]),"schedule":pd.DataFrame(columns=self.app.SCHEDULE),"risk":pd.DataFrame(columns=self.app.RISK)}
        portfolio=self.app.build_portfolio_summary(projects,lambda project,dataset,cols:data[dataset],self.app.date(2026,11,15))
        row=portfolio.iloc[0]
        self.assertTrue(pd.isna(row["Onaylı Bütçe"]))
        self.assertTrue(pd.isna(row["Onaylı Gerçekleşen"]))
        self.assertEqual(int(row["Kur Eksik Bütçe"]),1)
        self.assertEqual(int(row["Kur Eksik Gerçekleşen"]),1)
        self.assertEqual(row["Durum"],"VERİ EKSİK")

    def test_company_portfolio_does_not_hide_missing_actual_dates(self):
        projects=pd.DataFrame([{"_id":"prj-try","Proje Adı":"TRY Project","Para Birimi":"TRY"}])
        plan={c:None for c in self.app.PLAN}
        plan.update({"_id":"p1","WBS":"A.01","İş Kalemi":"Grobeton","Birim":"m³","Miktar":1,
                     "Birim Fiyat":100,"Para Birimi":"TRY","Kur":1,"Metraj Durumu":"Onaylandı"})
        actual={c:None for c in self.app.ACTUAL}
        actual.update({"_id":"a1","WBS":"A.01","İş Kalemi":"Grobeton","Birim":"m³","Gerçekleşen Miktar":1,
                       "Gerçekleşen Birim Fiyat":50,"Para Birimi":"TRY","Kur":1,"Tarih":None,"Durum":"Onaylandı"})
        data={"plan":pd.DataFrame([plan]),"actual":pd.DataFrame([actual]),"schedule":pd.DataFrame(columns=self.app.SCHEDULE),"risk":pd.DataFrame(columns=self.app.RISK)}
        portfolio=self.app.build_portfolio_summary(projects,lambda project,dataset,cols:data[dataset],self.app.date(2026,11,15))
        row=portfolio.iloc[0]
        self.assertTrue(pd.isna(row["Onaylı Gerçekleşen"]))
        self.assertEqual(int(row["Tarihi Eksik Gerçekleşen"]),1)
        self.assertEqual(row["Durum"],"VERİ EKSİK")

    def test_portfolio_does_not_publish_partial_currency_totals(self):
        projects=pd.DataFrame([{"_id":"prj-try","Proje Adı":"TRY Project","Para Birimi":"TRY"}])
        plans=[]; actuals=[]
        for rid,cur,rate in [("p1","TRY",1),("p2","EUR",None)]:
            row={c:None for c in self.app.PLAN}
            row.update({"_id":rid,"WBS":rid,"İş Kalemi":"Grobeton","Birim":"m³","Miktar":1,"Birim Fiyat":100,
                        "Para Birimi":cur,"Kur":rate,"Metraj Durumu":"Onaylandı"})
            plans.append(row)
        for rid,cur,rate in [("a1","TRY",1),("a2","EUR",None)]:
            row={c:None for c in self.app.ACTUAL}
            row.update({"_id":rid,"WBS":rid,"İş Kalemi":"Grobeton","Birim":"m³","Gerçekleşen Miktar":1,
                        "Gerçekleşen Birim Fiyat":50,"Para Birimi":cur,"Kur":rate,"Tarih":"2026-11-01","Durum":"Onaylandı"})
            actuals.append(row)
        data={"plan":pd.DataFrame(plans),"actual":pd.DataFrame(actuals),"schedule":pd.DataFrame(columns=self.app.SCHEDULE),"risk":pd.DataFrame(columns=self.app.RISK)}
        portfolio=self.app.build_portfolio_summary(projects,lambda project,dataset,cols:data[dataset],self.app.date(2026,11,15))
        row=portfolio.iloc[0]
        self.assertTrue(pd.isna(row["Onaylı Bütçe"]))
        self.assertTrue(pd.isna(row["Onaylı Gerçekleşen"]))
        self.assertEqual(int(row["Kur Eksik Bütçe"]),1)
        self.assertEqual(int(row["Kur Eksik Gerçekleşen"]),1)

    def _forecast_purchase(self, record_id="po1", amount=1000.0):
        row={c:None for c in self.app.PURCHASE}
        row.update({"_id":record_id,"Sipariş No":record_id,"WBS":"A.01","Malzeme / Hizmet":"Grobeton","Tedarikçi":"Supplier A",
                   "Sipariş Miktarı":1,"Birim":"m³","Birim Fiyat":amount,"Para Birimi":"TRY","Kur":1,
                   "Sipariş Tutarı":amount,"Faturalanan Tutar":400,"Faturalanan Miktar":0.4,"Fatura No":"INV-001",
                   "Fatura Tarihi":"2026-11-01","Durum":"Sipariş Verildi","Onay Durumu":"Onaylandı"})
        return row

    def _forecast_actual(self, record_id="ac1", amount=400.0, invoice="INV-001", date_value="2026-11-05", status="Onaylandı"):
        row={c:None for c in self.app.ACTUAL}
        row.update({"_id":record_id,"WBS":"A.01","İş Kalemi":"Grobeton","Birim":"m³","Gerçekleşen Miktar":0.4,
                    "Gerçekleşen Birim Fiyat":amount/0.4,"Gerçekleşen Tutar":amount,"Para Birimi":"TRY","Kur":1,
                    "Tarih":date_value,"Tedarikçi":"Supplier A","Belge No":invoice,"Durum":status,"Kayıt Türü":"Fatura"})
        return row

    def test_purchase_invoice_does_not_reduce_commitment_before_actual_approval(self):
        po=pd.DataFrame([self._forecast_purchase()])
        empty_contracts=pd.DataFrame(columns=self.app.CONTRACT); empty_changes=pd.DataFrame(columns=self.app.CHANGE)
        result=self.app.profit_forecast(pd.DataFrame(columns=self.app.PLAN),pd.DataFrame(columns=self.app.ACTUAL),po,
                                        empty_contracts,empty_changes,{"AC":0.0,"EV":0.0,"CPI":float("nan")},"TRY",self.app.date(2026,11,15))
        self.assertEqual(result["Açık Sipariş Taahhüdü"],1000.0)
        self.assertEqual(result["Eşleşmeyen Faturalı Sipariş"],1)

    def test_approved_invoice_actual_reduces_only_its_matching_commitment(self):
        po=pd.DataFrame([self._forecast_purchase()])
        actual=pd.DataFrame([self._forecast_actual()])
        empty_contracts=pd.DataFrame(columns=self.app.CONTRACT); empty_changes=pd.DataFrame(columns=self.app.CHANGE)
        result=self.app.profit_forecast(pd.DataFrame(columns=self.app.PLAN),actual,po,
                                        empty_contracts,empty_changes,{"AC":400.0,"EV":0.0,"CPI":float("nan")},"TRY",self.app.date(2026,11,15))
        self.assertEqual(result["Açık Sipariş Taahhüdü"],600.0)
        self.assertEqual(result["Onaylı Gerçekleşen"],400.0)
        self.assertEqual(result["Eşleşmeyen Faturalı Sipariş"],0)

    def test_future_invoice_actual_does_not_reduce_current_commitment(self):
        po=pd.DataFrame([self._forecast_purchase()])
        actual=pd.DataFrame([self._forecast_actual(date_value="2026-11-16")])
        empty_contracts=pd.DataFrame(columns=self.app.CONTRACT); empty_changes=pd.DataFrame(columns=self.app.CHANGE)
        result=self.app.profit_forecast(pd.DataFrame(columns=self.app.PLAN),actual,po,
                                        empty_contracts,empty_changes,{"AC":0.0,"EV":0.0,"CPI":float("nan")},"TRY",self.app.date(2026,11,15))
        self.assertEqual(result["Açık Sipariş Taahhüdü"],1000.0)
        self.assertEqual(result["Eşleşmeyen Faturalı Sipariş"],1)

    def test_ambiguous_purchase_invoice_key_is_not_double_matched(self):
        po=pd.DataFrame([self._forecast_purchase("po1",500.0),self._forecast_purchase("po2",500.0)])
        actual=pd.DataFrame([self._forecast_actual()])
        empty_contracts=pd.DataFrame(columns=self.app.CONTRACT); empty_changes=pd.DataFrame(columns=self.app.CHANGE)
        result=self.app.profit_forecast(pd.DataFrame(columns=self.app.PLAN),actual,po,
                                        empty_contracts,empty_changes,{"AC":400.0,"EV":0.0,"CPI":float("nan")},"TRY",self.app.date(2026,11,15))
        self.assertEqual(result["Açık Sipariş Taahhüdü"],1000.0)
        self.assertEqual(result["Eşleşmeyen Faturalı Sipariş"],2)

    def test_portfolio_page_is_limited_to_authorized_roles(self):
        self.assertIn("Şirket Portföyü",self.app.ROLE_PAGES["Yönetici"])
        self.assertIn("Şirket Portföyü",self.app.ROLE_PAGES["Proje Müdürü"])
        self.assertIn("Şirket Portföyü",self.app.ROLE_PAGES["Teknik Ofis"])
        self.assertIn("Şirket Portföyü",self.app.ROLE_PAGES["Finans / Muhasebe"])
        self.assertNotIn("Şirket Portföyü",self.app.ROLE_PAGES["Görüntüleyici"])

    def test_production_clamav_probe_scans_a_temporary_clean_file(self):
        import types
        executable=Path(sys.executable)
        original_run=self.app.subprocess.run
        calls=[]
        def fake_run(args,**kwargs):
            calls.append((args,kwargs))
            self.assertEqual(args[0],str(executable))
            self.assertIn("--no-summary",args)
            self.assertTrue(Path(args[-1]).is_file())
            return types.SimpleNamespace(returncode=0,stdout="clean",stderr="")
        self.app.subprocess.run=fake_run
        try:
            result=self.app.verify_clamav(str(executable))
            self.assertIn("temiz test dosyası",result)
            self.assertEqual(len(calls),1)
        finally:
            self.app.subprocess.run=original_run

    def test_production_clamav_probe_fails_closed_on_scanner_database_error(self):
        import types
        executable=Path(sys.executable)
        original_run=self.app.subprocess.run
        self.app.subprocess.run=lambda *args,**kwargs: types.SimpleNamespace(returncode=2,stdout="",stderr="No supported database files found")
        try:
            with self.assertRaisesRegex(RuntimeError,"No supported database files"):
                self.app.verify_clamav(str(executable))
        finally:
            self.app.subprocess.run=original_run

    def test_upload_rejects_non_excel_zip_and_zip_bomb_ratio(self):
        class Upload:
            def __init__(self, name, data): self.name=name; self._data=data
            def getvalue(self): return self._data
        with self.assertRaisesRegex(ValueError, "Excel OOXML arşivi bozuk|geçerli bir Excel"):
            self.app.validate_upload(Upload("fake.xlsx", b"PK\x03\x04not-an-excel"))
        # A tiny highly-compressible OOXML member is rejected before pandas/openpyxl reads it.
        import io
        buf=io.BytesIO()
        with zipfile.ZipFile(buf,"w",compression=zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", "x"*2_000_000)
            z.writestr("xl/workbook.xml", "<workbook/>")
        with self.assertRaisesRegex(ValueError, "sıkıştırma oranı"):
            self.app.validate_upload(Upload("bomb.xlsx",buf.getvalue()))

    def test_upload_validates_json_and_rejects_malformed_or_non_object_payloads(self):
        class Upload:
            def __init__(self, name, data): self.name=name; self.data=data
            def getvalue(self): return self.data
        original_scan=self.app.clamav_scan
        scanned=[]
        self.app.clamav_scan=lambda data,suffix: scanned.append(suffix) or "TEST SCAN"
        try:
            name,ext,data,digest=self.app.validate_upload(Upload("C:\\fakepath\\backup.json", b'{"app":"emir-asel"}'))
            self.assertEqual(name,"backup.json")
            self.assertEqual(ext,".json")
            self.assertEqual(digest,self.app.hashlib.sha256(data).hexdigest())
            self.assertIn(".json",scanned)
            with self.assertRaisesRegex(ValueError,"geçerli UTF-8 JSON"):
                self.app.validate_upload(Upload("bad.json", b"{broken"))
            with self.assertRaisesRegex(ValueError,"en üst seviyesi nesne"):
                self.app.validate_upload(Upload("array.json", b"[]"))
        finally:
            self.app.clamav_scan=original_scan

    def test_upload_validates_csv_encoding_and_runs_scanner(self):
        class Upload:
            def __init__(self, name, data): self.name=name; self.data=data
            def getvalue(self): return self.data
        original_scan=self.app.clamav_scan
        scanned=[]
        self.app.clamav_scan=lambda data,suffix: scanned.append((suffix,len(data))) or "TEST SCAN"
        try:
            tr_bytes="İş Kalemi,Miktar\nBeton,12".encode("cp1254")
            _name,ext,data,_digest=self.app.validate_upload(Upload("metraj.csv",tr_bytes))
            self.assertEqual(ext,".csv")
            self.assertEqual(self.app.read_csv_bytes(data).shape,(1,2))
            self.assertEqual(scanned[0][0],".csv")
            with self.assertRaisesRegex(ValueError,"ikili içerik"):
                self.app.validate_upload(Upload("binary.csv",b"a,b\n1,2\x00"))
        finally:
            self.app.clamav_scan=original_scan

    def test_upload_rejects_macro_enabled_xlsm(self):
        class Upload:
            name="macro.xlsm"
            def __init__(self,data): self.data=data
            def getvalue(self): return self.data
        import io
        buf=io.BytesIO()
        with zipfile.ZipFile(buf,"w",compression=zipfile.ZIP_STORED) as z:
            z.writestr("[Content_Types].xml","<Types/>")
            z.writestr("xl/workbook.xml","<workbook/>")
            z.writestr("xl/vbaProject.bin",b"MZ\x00macro")
        with self.assertRaisesRegex(ValueError, "Makro içeren XLSM"):
            self.app.validate_upload(Upload(buf.getvalue()))

    def test_ai_takeoff_quantity_requires_page_evidence_and_known_unit(self):
        payload={"items":[
            {"page_no":1,"discipline":"Statik","item_description":"C30/37 radye temel betonu","technical_specification":"C30/37","quantity":42.5,"unit":"m³","measurement_basis":"dimension_calculation","calculation":"8.5 m × 5.0 m × 1.0 m = 42.5 m3","source_evidence":"Pafta S-01: temel 8.50 × 5.00 m; kalınlık 1.00 m","confidence":0.91},
            {"page_no":1,"discipline":"Statik","item_description":"Kolon betonu","technical_specification":"C30/37","quantity":999,"unit":"m3","measurement_basis":"dimension_calculation","calculation":"model tahmini","source_evidence":"","confidence":0.99},
            {"page_no":9,"discipline":"Statik","item_description":"Temel betonu","quantity":12,"unit":"m3","measurement_basis":"direct_table","source_evidence":"Poz tablosu 12 m3","confidence":0.99},
            {"page_no":1,"discipline":"Mimari","item_description":"Duvar kaplama","quantity":25,"unit":"belirsiz","measurement_basis":"direct_table","source_evidence":"Pafta notu 25","confidence":0.99},
        ]}
        frame=self.app.validate_ai_takeoff_payload(payload,"test.pdf","abc123",[1])
        self.assertEqual(len(frame),3)
        confirmed=frame[frame["İş Kalemi"].eq("C30/37 radye temel betonu")].iloc[0]
        self.assertAlmostEqual(float(confirmed["Miktar"]),42.5)
        self.assertEqual(confirmed["Birim"],"m3")
        unsubstantiated=frame[frame["İş Kalemi"].eq("Kolon betonu")].iloc[0]
        self.assertTrue(pd.isna(unsubstantiated["Miktar"]))
        invalid_unit=frame[frame["İş Kalemi"].eq("Duvar kaplama")].iloc[0]
        self.assertTrue(pd.isna(invalid_unit["Miktar"]))

    def test_dimension_calculation_is_recomputed_and_mismatch_quantity_is_removed(self):
        self.assertTrue(self.app.calculation_matches_quantity("8.5 m × 5.0 m × 1.0 m = 42.5 m3",42.5))
        self.assertFalse(self.app.calculation_matches_quantity("8.5 m × 5.0 m × 1.0 m = 42.5 m3",45.0))
        self.assertTrue(self.app.calculation_matches_quantity("8,5m × 5m × 1m = 42,5 m³",42.5))
        payload={"items":[
            {"page_no":1,"item_description":"Radye beton","unit":"m3","quantity":45,"measurement_basis":"dimension_calculation","calculation":"8.5 × 5 × 1 = 42.5 m3","source_evidence":"Pafta S-01 ölçüsü 8.5 x 5 x 1 m","confidence":0.95},
            {"page_no":1,"item_description":"Doğru radye beton","unit":"m3","quantity":42.5,"measurement_basis":"dimension_calculation","calculation":"8.5 × 5 × 1 = 42.5 m3","source_evidence":"Pafta S-01 ölçüsü 8.5 x 5 x 1 m","confidence":0.95},
        ]}
        rows=self.app.validate_ai_takeoff_payload(payload,"test.pdf","hash",[1])
        wrong=rows[rows["İş Kalemi"].eq("Radye beton")].iloc[0]
        right=rows[rows["İş Kalemi"].eq("Doğru radye beton")].iloc[0]
        self.assertTrue(pd.isna(wrong["Miktar"]))
        self.assertAlmostEqual(float(right["Miktar"]),42.5)

    def test_ai_pdf_takeoff_renders_pages_and_parses_mocked_response(self):
        import fitz, json
        from types import SimpleNamespace
        from unittest.mock import patch
        doc=fitz.open(); page=doc.new_page(width=595,height=842); page.insert_text((50,80),"S-01 TEMEL PLANI 8.50 x 5.00 m; kalinlik 1.00 m")
        pdf_bytes=doc.tobytes(); doc.close()
        output={"items":[{"page_no":1,"discipline":"Statik","drawing_no":"S-01","floor_or_location":"Temel","item_description":"C30/37 radye temel betonu","technical_specification":"C30/37","quantity":42.5,"unit":"m3","measurement_basis":"dimension_calculation","calculation":"8.50 × 5.00 × 1.00 = 42.50 m3","source_evidence":"S-01 pafta notu: 8.50 x 5.00 m; kalınlık 1.00 m","confidence":0.9}]}
        class Responses:
            def create(self,**kwargs):
                self.last=kwargs
                return SimpleNamespace(output_text=json.dumps(output,ensure_ascii=False))
        fake_client=SimpleNamespace(responses=Responses())
        with patch.object(self.app,"ai_client",return_value=(fake_client,"mock-model")):
            result=self.app.extract_pdf_project_takeoff(pdf_bytes,"sample.pdf","deadbeef","test project",1,1)
        self.assertEqual(len(result),1)
        self.assertEqual(int(result.iloc[0]["Sayfa"]),1)
        self.assertAlmostEqual(float(result.iloc[0]["Miktar"]),42.5)
        self.assertEqual(result.iloc[0]["Kaynak SHA256"],"deadbeef")

    def test_pdf_takeoff_to_pending_budget_end_to_end_keeps_price_acceptance_explicit(self):
        import fitz, json
        from types import SimpleNamespace
        from unittest.mock import patch
        doc=fitz.open(); page=doc.new_page(width=595,height=842); page.insert_text((40,70),"S-02 Radye temel 8 x 5 x 1 m")
        data=doc.tobytes(); doc.close()
        response={"items":[{"page_no":1,"discipline":"Statik","drawing_no":"S-02","floor_or_location":"Temel","item_description":"C30/37 radye beton","technical_specification":"C30/37","quantity":40,"unit":"m3","measurement_basis":"dimension_calculation","calculation":"8 x 5 x 1 = 40 m3","source_evidence":"S-02: 8 x 5 m; temel kalınlığı 1 m","confidence":0.93}]}
        class Responses:
            def create(self,**kwargs): return SimpleNamespace(output_text=json.dumps(response,ensure_ascii=False))
        client=SimpleNamespace(responses=Responses())
        with patch.object(self.app,"ai_client",return_value=(client,"vision-model")):
            candidates=self.app.extract_pdf_project_takeoff(data,"S-02.pdf","sha256-test","villa statik temel",1,1)
        self.assertEqual(len(candidates),1)
        self.assertAlmostEqual(float(candidates.iloc[0]["Miktar"]),40.0)
        candidates.loc[0,"Bütçeye Aktar"]=True
        pending=self.app.build_ai_takeoff_plan_rows(candidates,pd.DataFrame(columns=self.app.PLAN),"TRY")
        self.assertEqual(len(pending),1)
        self.assertTrue(pd.isna(pending.iloc[0]["Birim Fiyat"]))
        candidates.loc[0,"Önerilen Fiyat"]=3200.0
        candidates.loc[0,"Önerilen Fiyat Para Birimi"]="TRY"
        candidates.loc[0,"Önerilen Fiyat Kaynağı"]="YFK 2026 Resmi Poz"
        candidates.loc[0,"Önerilen Fiyat Yılı"]=2026
        candidates.loc[0,"Fiyat Eşleşme Durumu"]="YÜKSEK EŞLEŞME — ONAY GEREKLİ"
        candidates.loc[0,"Yerel Fiyatı Kabul Et"]=True
        priced=self.app.build_ai_takeoff_plan_rows(candidates,pd.DataFrame(columns=self.app.PLAN),"TRY")
        self.assertAlmostEqual(float(priced.iloc[0]["Birim Fiyat"]),3200.0)
        self.assertEqual(priced.iloc[0]["Metraj Durumu"],"Kontrol Bekliyor")

    def test_tavily_search_sends_country_and_rejects_unsafe_result_urls(self):
        import os
        from unittest.mock import patch
        class Response:
            def raise_for_status(self): pass
            def json(self):
                return {"results":[
                    {"title":"Supplier product price","url":"https://supplier.example/product","content":"3,450 TL/m3","score":0.91},
                    {"title":"Unsafe local","url":"http://127.0.0.1/admin","content":"price 5 TL","score":1.0},
                    {"title":"Invalid scheme","url":"javascript:alert(1)","content":"price 5 TL","score":1.0},
                ]}
        calls=[]
        def fake_post(url,**kwargs): calls.append((url,kwargs)); return Response()
        old=os.environ.get("TAVILY_API_KEY")
        os.environ["TAVILY_API_KEY"]="test-token"
        try:
            with patch("requests.post",side_effect=fake_post):
                rows=self.app.tavily_search_web("C30/37 hazır beton m3 fiyat Adana",6)
            self.assertEqual(len(rows),1)
            self.assertEqual(rows[0]["url"],"https://supplier.example/product")
            self.assertEqual(calls[0][0],self.app.MARKET_PRICE_API_URL)
            self.assertEqual(calls[0][1]["json"]["country"],"turkey")
            self.assertTrue(calls[0][1]["json"]["include_raw_content"])
        finally:
            if old is None: os.environ.pop("TAVILY_API_KEY",None)
            else: os.environ["TAVILY_API_KEY"]=old

    def test_market_price_is_kept_only_when_price_excerpt_exists_in_source(self):
        import json
        from types import SimpleNamespace
        from unittest.mock import patch
        results=[{"source_id":"S1","title":"Supplier price","url":"https://supplier.example/price","content":"C30/37 hazır beton güncel fiyatı 3.450 TL/m3 KDV hariç.","score":0.9,"published_date":"2026-11-01"}]
        output={"quotes":[{"source_id":"S1","price":3450,"currency":"TRY","source_unit":"m3","price_text":"3.450 TL/m3","product_name":"C30/37 hazır beton","technical_match":"YES","vat_status":"excluded","shipping_status":"unknown","confidence":0.9}]}
        class Responses:
            def create(self,**kwargs): return SimpleNamespace(output_text=json.dumps(output,ensure_ascii=False))
        client=SimpleNamespace(responses=Responses())
        import os
        prior=os.environ.get("OPENAI_API_KEY")
        prior_model=os.environ.get("OPENAI_MODEL")
        os.environ["OPENAI_API_KEY"]="test-token"
        os.environ["OPENAI_MODEL"]="mock-model"
        try:
            with patch.object(self.app,"ai_client",return_value=(client,"mock-model")):
                rows=self.app.extract_market_quotes_ai("C30/37 beton","C30/37", "m3","Adana",results)
            self.assertEqual(float(rows.iloc[0]["Fiyat"]),3450.0)
            self.assertEqual(rows.iloc[0]["Durum"],"FİYAT ADAYI — UZMAN KONTROLÜ GEREKLİ")
            output["quotes"][0]["price_text"]="4.999 TL/m3"
            with patch.object(self.app,"ai_client",return_value=(client,"mock-model")):
                rows=self.app.extract_market_quotes_ai("C30/37 beton","C30/37", "m3","Adana",results)
            self.assertTrue(pd.isna(rows.iloc[0]["Fiyat"]))
            self.assertEqual(rows.iloc[0]["Durum"],"KAYNAK BULUNDU — AÇIK FİYAT OKUNAMADI")
        finally:
            if prior is None: os.environ.pop("OPENAI_API_KEY",None)
            else: os.environ["OPENAI_API_KEY"]=prior
            if prior_model is None: os.environ.pop("OPENAI_MODEL",None)
            else: os.environ["OPENAI_MODEL"]=prior_model

    def test_market_quote_requires_manual_technical_acceptance_and_converts_unit_price(self):
        quote={"Teknik Uyum":"Kontrol Bekliyor","Kaynak URL":"https://supplier.example/price","Fiyat":38.5,"Para Birimi":"TRY","Kaynak Birimi":"kg","Hedef Birim":"ton","İş Kalemi":"B420C nervürlü donatı","Teknik Tarif":"B420C Ø12-Ø32","Kaynak Başlığı":"Üretici fiyat sayfası","Yayın Tarihi":"2026-11-01","Güven Skoru":0.85}
        with self.assertRaisesRegex(ValueError,"Teknik Uyum"):
            self.app.market_quote_to_library_row(quote,"buyer01")
        quote["Teknik Uyum"]="Uygun"
        row=self.app.market_quote_to_library_row(quote,"buyer01")
        self.assertAlmostEqual(row["Birim Fiyat"],38500.0)
        self.assertEqual(row["Birim"],"ton")
        self.assertEqual(row["Kaynak URL"],quote["Kaynak URL"])
        self.assertEqual(row["Teknik Özellik"],quote["Teknik Tarif"])

    def test_ai_takeoff_transfer_never_accepts_unsubstantiated_or_low_confidence_quantity(self):
        candidate={c:None for c in self.app.TAKEOFF_CANDIDATES}
        candidate.update({"_id":"takeoff-1","Kaynak SHA256":"hash","Kaynak Dosya":"S-01.pdf","Sayfa":1,"İş Kalemi":"Grobeton","Teknik Tarif":"C16/20","Birim":"m3","Miktar":20,"Ölçüm Yöntemi":"dimension_calculation","Hesap Formülü":"5 x 4 x 1 = 20 m3","Kanıt":"S-01: 5 x 4 x 1 m","Güven Skoru":0.9,"Bütçeye Aktar":True,"Aktarıldı":False})
        rows=self.app.build_ai_takeoff_plan_rows(pd.DataFrame([candidate]),pd.DataFrame(columns=self.app.PLAN),"TRY")
        self.assertEqual(len(rows),1)
        self.assertTrue(str(rows.iloc[0]["WBS"]).startswith("MET-AI-"))
        self.assertTrue(pd.isna(rows.iloc[0]["Birim Fiyat"]))
        candidate["Önerilen Fiyat"]=2500.0
        candidate["Önerilen Fiyat Para Birimi"]="TRY"
        candidate["Önerilen Fiyat Kaynağı"]="YFK 2026/11 Resmi Rayiç"
        candidate["Önerilen Fiyat Yılı"]=2026
        candidate["Fiyat Eşleşme Durumu"]="YÜKSEK EŞLEŞME — ONAY GEREKLİ"
        candidate["Yerel Fiyatı Kabul Et"]=True
        priced=self.app.build_ai_takeoff_plan_rows(pd.DataFrame([candidate]),pd.DataFrame(columns=self.app.PLAN),"TRY")
        self.assertAlmostEqual(float(priced.iloc[0]["Birim Fiyat"]),2500.0)
        self.assertEqual(priced.iloc[0]["Fiyat Kaynağı"],"YFK 2026/11 Resmi Rayiç")
        candidate["Fiyat Eşleşme Durumu"]="ESKİ FİYAT — GÜNCEL KAYNAK KONTROLÜ GEREKLİ"
        with self.assertRaisesRegex(ValueError,"eski fiyat"):
            self.app.build_ai_takeoff_plan_rows(pd.DataFrame([candidate]),pd.DataFrame(columns=self.app.PLAN),"TRY")
        candidate["Yerel Fiyatı Kabul Et"]=False
        candidate["Güven Skoru"]=0.4
        with self.assertRaisesRegex(ValueError,"%60 altında"):
            self.app.build_ai_takeoff_plan_rows(pd.DataFrame([candidate]),pd.DataFrame(columns=self.app.PLAN),"TRY")
        candidate["Güven Skoru"]=0.9; candidate["Kanıt"]=""
        with self.assertRaisesRegex(ValueError,"kanıtı yok"):
            self.app.build_ai_takeoff_plan_rows(pd.DataFrame([candidate]),pd.DataFrame(columns=self.app.PLAN),"TRY")

    def test_takeoff_and_market_quote_persistence_round_trip(self):
        project="AI Takeoff Persistence"
        takeoff={c:None for c in self.app.TAKEOFF_CANDIDATES}; takeoff.update({"_id":"takeoff-persist","Sayfa":3,"İş Kalemi":"Temel membranı","Birim":"m2","Miktar":110,"Kanıt":"Pafta A-04"})
        quote={c:None for c in self.app.MARKET_QUOTES}; quote.update({"_id":"quote-persist","İş Kalemi":"Temel membranı 3.5 mm","Fiyat":420,"Para Birimi":"TRY","Kaynak URL":"https://supplier.example/membrane"})
        self.app.save_datasets_atomic(project,{"takeoff_candidates":(pd.DataFrame([takeoff]),self.app.TAKEOFF_CANDIDATES),"market_quotes":(pd.DataFrame([quote]),self.app.MARKET_QUOTES)})
        self.assertEqual(self.app.load_ds(project,"takeoff_candidates",self.app.TAKEOFF_CANDIDATES).iloc[0]["İş Kalemi"],"Temel membranı")
        self.assertAlmostEqual(float(self.app.load_ds(project,"market_quotes",self.app.MARKET_QUOTES).iloc[0]["Fiyat"]),420)

    def test_manual_dxf_layer_mapping_prefers_hatch_and_closed_polylines(self):
        detail=pd.DataFrame([
            {"Tip":"HATCH","Katman":"LAYER_MEMBRANE","Kaynak":"MODEL","Uzunluk (m)":None,"Alan (m²)":110.0},
            {"Tip":"LWPOLYLINE","Katman":"LAYER_MEMBRANE","Kaynak":"MODEL","Uzunluk (m)":42.0,"Alan (m²)":110.0},
            {"Tip":"LWPOLYLINE","Katman":"PIPE_A","Kaynak":"MODEL","Uzunluk (m)":28.5,"Alan (m²)":None},
            {"Tip":"LINE","Katman":"PIPE_A","Kaynak":"MODEL","Uzunluk (m)":28.5,"Alan (m²)":None},
            {"Tip":"INSERT","Katman":"VALVE_SYMBOLS","Kaynak":"MODEL","Uzunluk (m)":None,"Alan (m²)":None},
            {"Tip":"INSERT","Katman":"VALVE_SYMBOLS","Kaynak":"MODEL","Uzunluk (m)":None,"Alan (m²)":None},
        ])
        result={"factor":1.0,"detail":detail.to_dict("records")}
        mappings=pd.DataFrame([
            {"Katman":"LAYER_MEMBRANE","İmalat":"Temel membranı 3,5 mm SBS","Ölçüm Türü":"Alan (m²)","WBS":"IZL.01"},
            {"Katman":"PIPE_A","İmalat":"Temiz su borusu DN25","Ölçüm Türü":"Uzunluk (m)","WBS":"MEK.01"},
            {"Katman":"VALVE_SYMBOLS","İmalat":"Vana","Ölçüm Türü":"Adet","WBS":"MEK.02"},
        ])
        rows=self.app.dxf_manual_layer_takeoff_candidates(result,mappings,pd.DataFrame(),"TRY","test.dxf","sha")
        self.assertEqual(len(rows),3)
        by_layer=rows.set_index("Kaynak Layer")
        self.assertAlmostEqual(float(by_layer.loc["LAYER_MEMBRANE","Miktar"]),110.0)
        self.assertAlmostEqual(float(by_layer.loc["PIPE_A","Miktar"]),28.5)
        self.assertEqual(float(by_layer.loc["VALVE_SYMBOLS","Miktar"]),2.0)
        self.assertEqual(by_layer.loc["PIPE_A","Birim"],"m")
        self.assertEqual(by_layer.loc["VALVE_SYMBOLS","Birim"],"adet")

    def test_manual_dxf_mapping_requires_known_drawing_units(self):
        result={"factor":None,"detail":[{"Tip":"HATCH","Katman":"A","Kaynak":"MODEL","Alan (m²)":100,"Uzunluk (m)":None}]}
        mappings=pd.DataFrame([{"Katman":"A","İmalat":"Temel yalıtımı","Ölçüm Türü":"Alan (m²)"}])
        rows=self.app.dxf_manual_layer_takeoff_candidates(result,mappings,pd.DataFrame(),"TRY")
        self.assertTrue(rows.empty)

    def test_concurrent_change_is_not_overwritten(self):
        project = "Conflict Project"
        frame = pd.DataFrame([self._plan_row("plan-conflict")])
        self.app.save_ds(project, "plan", frame, self.app.PLAN)
        changed = dict(self._plan_row("plan-conflict"))
        changed["Miktar"] = 21.0
        with self.app.ENGINE.begin() as connection:
            from sqlalchemy import text
            connection.execute(
                text("UPDATE ea_records SET payload=:payload WHERE project_key=:project AND dataset='plan' AND record_id='plan-conflict'"),
                {"payload": self.app.payload(changed), "project": project},
            )
        frame.loc[0, "Miktar"] = 22.0
        with self.assertRaisesRegex(RuntimeError, "başka bir kullanıcı tarafından değiştirilmiş"):
            self.app.save_ds(project, "plan", frame, self.app.PLAN)


    def test_approved_business_records_reopen_after_any_material_edit(self):
        cases = [
            ("plan", "Metraj Durumu", "Miktar", 10.0, None, None),
            ("actual", "Durum", "Gerçekleşen Miktar", 10.0, "Onay Tarihi", "Onaylayan"),
            ("purchase", "Onay Durumu", "Birim Fiyat", 100.0, None, None),
            ("contracts", "Durum", "Orijinal Bedel", 100000.0, "Onay Tarihi", "Onaylayan"),
            ("changes", "Durum", "Maliyet Etkisi", 25000.0, "Onay Tarihi", "Onaylayan"),
            ("hakediş", "Durum", "Brüt Tutar", 50000.0, None, None),
            ("employer_hakediş", "Durum", "Brüt İmalat Tutarı", 75000.0, "Onay Tarihi", "Onaylayan"),
        ]
        for dataset, status_col, edit_col, original, date_col, approver_col in cases:
            with self.subTest(dataset=dataset):
                project = f"Approved Guard {dataset}"
                row = {c: None for c in self.app.DS[dataset]}
                row.update({"_id": f"approved-{dataset}", status_col: "Onaylandı", edit_col: original})
                if date_col: row[date_col] = "2026-11-15"
                if approver_col: row[approver_col] = "approver01"
                self.app.save_ds(project, dataset, pd.DataFrame([row]), self.app.DS[dataset])
                edited = self.app.load_ds(project, dataset, self.app.DS[dataset])
                edited.at[0, edit_col] = original + 1
                self.app.save_ds(project, dataset, edited, self.app.DS[dataset])
                saved = self.app.load_ds(project, dataset, self.app.DS[dataset]).iloc[0]
                self.assertEqual(saved[status_col], "Kontrol Bekliyor", f"{dataset} onayı yeniden açılmadı")
                if date_col: self.assertTrue(self.app.missing(saved[date_col]), f"{dataset} onay tarihi temizlenmedi")
                if approver_col: self.assertTrue(self.app.missing(saved[approver_col]), f"{dataset} onaylayanı temizlenmedi")
                note_col = "Not" if dataset == "purchase" else "Açıklama"
                self.assertIn("yeniden kontrol ve onay zorunlu", str(saved.get(note_col, "")))

    def test_approved_business_records_cannot_be_deleted(self):
        cases = [
            ("plan", "Metraj Durumu"), ("actual", "Durum"), ("purchase", "Onay Durumu"),
            ("contracts", "Durum"), ("changes", "Durum"), ("hakediş", "Durum"),
            ("employer_hakediş", "Durum"),
        ]
        for dataset, status_col in cases:
            with self.subTest(dataset=dataset):
                project = f"Approved Delete Guard {dataset}"
                row = {c: None for c in self.app.DS[dataset]}
                row.update({"_id": f"keep-{dataset}", status_col: "Onaylandı"})
                self.app.save_ds(project, dataset, pd.DataFrame([row]), self.app.DS[dataset])
                with self.assertRaisesRegex(ValueError, "Onaylanmış .* kaydı silinemez"):
                    self.app.save_ds(project, dataset, pd.DataFrame(columns=self.app.DS[dataset]), self.app.DS[dataset])
                self.assertEqual(len(self.app.load_ds(project, dataset, self.app.DS[dataset])), 1)

    def test_approved_rows_stay_approved_when_business_fields_are_unchanged(self):
        project = "Approved Noop Guard"
        row = {c: None for c in self.app.ACTUAL}
        row.update({"_id": "stable-approved", "Durum": "Onaylandı", "Gerçekleşen Miktar": 4.0,
                    "Onay Tarihi": "2026-11-15", "Onaylayan": "approver01"})
        self.app.save_ds(project, "actual", pd.DataFrame([row]), self.app.ACTUAL)
        unchanged = self.app.load_ds(project, "actual", self.app.ACTUAL)
        self.app.save_ds(project, "actual", unchanged, self.app.ACTUAL)
        self.assertEqual(self.app.load_ds(project, "actual", self.app.ACTUAL).iloc[0]["Durum"], "Onaylandı")

    def test_takeoff_and_market_provenance_schema_round_trip(self):
        self.assertIn("takeoff_candidates", self.app.DATE_FIELDS)
        self.assertEqual(self.app.DATE_FIELDS["takeoff_candidates"], ["Yayın Tarihi", "Kontrol Tarihi"])
        self.assertIn("Kaynak Sayfa", self.app.TAKEOFF_CANDIDATES)
        self.assertIn("Kaynak SHA256", self.app.MARKET_QUOTES)
        self.assertIn("Kaynak SHA256", self.app.LIB)
        payload = {"items": [{"page_no": 2, "item_description": "Temel membranı", "technical_specification": "SBS 3,5 mm",
                              "quantity": 110, "unit": "m2", "measurement_basis": "direct_table",
                              "source_evidence": "Pafta M-02: 110 m²", "confidence": 0.92}]}
        takeoff = self.app.validate_ai_takeoff_payload(payload, "M-02.pdf", "source-hash", [2])
        self.assertEqual(takeoff.iloc[0]["Kaynak Sayfa"], 2)
        self.assertEqual(takeoff.iloc[0]["Sayfa"], 2)
        self.assertEqual(takeoff.iloc[0]["Kaynak SHA256"], "source-hash")
        self.assertFalse(self.app.missing(takeoff.iloc[0]["Kontrol Tarihi"]))

    def test_market_quote_persists_source_digest_file_and_query_without_fabricating_price(self):
        from unittest.mock import patch
        source = {"source_id": "S1", "title": "3,5 mm SBS Membran Ürün Sayfası", "url": "https://supplier.example/sbs",
                  "published_date": "2026-11-01", "content": "3,5 mm SBS membran rulo fiyatı 420 TRY/m2. Teknik ürün açıklaması." , "score": 0.9}
        with patch.object(self.app, "configured_providers", return_value=[]):
            rows = self.app.extract_market_quotes_ai("Temel membranı", "SBS 3,5 mm", "m2", "Adana", [source])
        self.assertEqual(len(rows), 1)
        row = rows.iloc[0]
        self.assertEqual(row["Kaynak Dosya"], "Web araştırması")
        self.assertEqual(row["Kaynak URL"], source["url"])
        self.assertTrue(len(str(row["Kaynak SHA256"])) == 64)
        self.assertIn("Temel membranı", row["Araştırma Sorgusu"])
        self.assertTrue(pd.isna(row["Fiyat"]), "AI fiyat çıkarımı yapılmadan fiyat uydurulmamalı")

    def test_packaged_dxf_fixture_parses_real_entities_without_errors(self):
        from pathlib import Path
        fixture = Path(__file__).resolve().parents[1] / "test_project" / "Villa_Test_Pafta.dxf"
        result = self.app.read_dxf(fixture)
        self.assertGreater(len(result["detail"]), 0)
        self.assertIsNotNone(result["factor"])
        self.assertEqual(result["errors"], [])


    def test_real_world_kalem_column_auto_maps_and_unverified_prices_are_held(self):
        raw = pd.DataFrame({
            "Kalem": ["Radye temel betonu", "Temel / perde donatısı", "Elektrik talep gücü"],
            "Kod": ["M01", "M02", "E01"],
            "Miktar": [None, 3.673, 9.776],
            "Birim": ["m³", "Ton", "kW"],
            "Birim Fiyat": [3700, 30500, None],
            "Kaynak Durumu": ["AÇIK", "DOĞRULANDI", "DOĞRULANDI"],
        })
        mapping = {key: self.app.find_column(raw, candidates) or "—"
                   for key, candidates in self.app.TAKEOFF_EXCEL_COLUMN_MAPS.items()}
        self.assertEqual(mapping["İş Kalemi"], "Kalem")
        self.app.BASE_CURRENCY = "TRY"
        candidates = self.app.build_excel_takeoff_rows(
            raw, mapping, "TR (1.234,56)", "TRY", pd.DataFrame(columns=self.app.LIB),
            "Villa_Insaat_Proje_Capraz_Metraj_v2.xlsx", "Metraj"
        )
        self.assertEqual(len(candidates), 3)
        self.assertTrue(candidates["İş Kalemi"].str.len().gt(2).all())
        self.assertTrue(pd.to_numeric(candidates["Birim Fiyat"], errors="coerce").isna().all())
        self.assertTrue(candidates["Açıklama"].str.contains("fiyat yılı ve kaynak", case=False).iloc[1])
        transferable = candidates[
            pd.to_numeric(candidates["Miktar"], errors="coerce").gt(0)
            & candidates["Birim"].fillna("").astype(str).str.strip().ne("")
        ]
        self.assertEqual(len(transferable), 2)

    def test_weighted_average_stock_moving_cost_and_negative_stock_warning(self):
        rows = []
        base = {column: None for column in self.app.INVENTORY}
        for rid, date_value, move, qty, price in [
            ("in-1", "2026-01-01", "Giriş", 10, 100),
            ("in-2", "2026-01-02", "Giriş", 10, 200),
            ("out-1", "2026-01-03", "Çıkış", 5, 0),
        ]:
            row = base.copy()
            row.update({"_id": rid, "Tarih": date_value, "Depo": "Ana Depo", "Malzeme Kodu": "DEMIR-12",
                        "Malzeme": "Nervürlü demir Ø12", "Hareket Türü": move, "Miktar": qty,
                        "Birim": "kg", "Birim Fiyat": price, "Para Birimi": "TRY", "Kur": 1})
            rows.append(row)
        summary, movements = self.app.weighted_average_stock(pd.DataFrame(rows), "TRY")
        self.assertEqual(len(summary), 1)
        self.assertAlmostEqual(float(summary.iloc[0]["Mevcut Stok"]), 15.0)
        self.assertAlmostEqual(float(summary.iloc[0]["Hareketli Ortalama Birim Maliyet"]), 150.0)
        self.assertAlmostEqual(float(summary.iloc[0]["Stok Değeri"]), 2250.0)
        self.assertEqual(int(summary.iloc[0]["Negatif Stok Uyarısı"]), 0)
        # A second outbound movement exceeding remaining stock must be flagged.
        extra = rows[-1].copy(); extra.update({"_id": "out-2", "Tarih": "2026-01-04", "Miktar": 20})
        flagged, _ = self.app.weighted_average_stock(pd.DataFrame(rows + [extra]), "TRY")
        self.assertGreaterEqual(int(flagged.iloc[0]["Negatif Stok Uyarısı"]), 1)
        self.assertEqual(flagged.iloc[0]["Değerleme Durumu"], "KONTROL GEREKLİ")

    def test_weighted_progress_uses_cost_weight_then_duration_fallback(self):
        weighted = pd.DataFrame({"Tamamlanma (%)": [20, 80], "Ağırlık (%)": [70, 30], "Süre (gün)": [5, 5]})
        self.assertAlmostEqual(self.app.weighted_progress(weighted), 38.0)
        fallback = pd.DataFrame({"Tamamlanma (%)": [20, 80], "Ağırlık (%)": [0, 0], "Süre (gün)": [1, 3]})
        self.assertAlmostEqual(self.app.weighted_progress(fallback), 65.0)
        self.assertEqual(self.app.weighted_progress(pd.DataFrame()), 0.0)

    def test_quality_checks_flag_incomplete_budget_and_schedule_data(self):
        plan = pd.DataFrame([self._plan_row("qc-1"), self._plan_row("qc-2")])
        plan.loc[1, "WBS"] = "KAB.01"
        plan.loc[1, "Poz No"] = plan.loc[0, "Poz No"]
        plan.loc[1, "Miktar"] = 0
        plan.loc[1, "Birim"] = ""
        plan.loc[1, "Para Birimi"] = ""
        plan.loc[1, "Başlangıç"] = "2026-12-20"
        plan.loc[1, "Bitiş"] = "2026-12-10"
        actual = pd.DataFrame([{column: None for column in self.app.ACTUAL}])
        actual.loc[0, "_id"] = "qc-actual"
        actual.loc[0, "Durum"] = "Kontrol Bekliyor"
        schedule = pd.DataFrame([{column: None for column in self.app.SCHEDULE}])
        schedule.loc[0, "_id"] = "qc-schedule"
        schedule.loc[0, "WBS"] = "KAB.01"
        schedule.loc[0, "Tamamlanma (%)"] = 120
        library = pd.DataFrame([{column: None for column in self.app.LIB}])
        checks = self.app.quality_checks(plan, actual, schedule, pd.DataFrame(columns=self.app.RISK), library, "TRY")
        counts = dict(zip(checks["Kontrol"], checks["Uyarı Sayısı"]))
        self.assertEqual(counts["Tekrarlı WBS/Poz"], 2)
        self.assertEqual(counts["Eksik/sıfır/negatif miktar"], 1)
        self.assertEqual(counts["Birim eksik"], 1)
        self.assertEqual(counts["Para birimi eksik"], 1)
        self.assertEqual(counts["Onaysız gerçekleşen"], 1)
        self.assertEqual(counts["İlerleme aralık dışı"], 1)
        self.assertEqual(counts["Fiyat kaynağı eksik"], 1)

if __name__ == "__main__":
    unittest.main(verbosity=2)
