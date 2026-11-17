"""Focused pure-logic regression tests for V15.2.0.
The Streamlit page shell is not imported; this isolates deterministic business rules.
"""
import ast, difflib, math, re, tempfile, unittest, uuid
from datetime import datetime
from pathlib import Path
import numpy as np
import pandas as pd

HERE = Path(__file__).resolve()
CANDIDATES = [HERE.with_name('app_V15_0_0.py'), HERE.parents[1] / 'app.py']
SOURCE_PATH = next((p for p in CANDIDATES if p.exists()), None)
if SOURCE_PATH is None: raise FileNotFoundError('V15 app.py source not found beside test or in release root')
SOURCE = SOURCE_PATH.read_text(encoding='utf-8')
TREE = ast.parse(SOURCE)
WANTED_FUNCS = {
    'missing','uid','norm','norm_col','cur_code','num','ensure_frame','categorize','aggregate_ifc_qto_candidates',
    'classify_work_item','apply_work_classification','unit_norm','convert_qty','work_scope_compatible','merge_takeoff_candidates',
    'tech_compatible','_description_similarity','price_match','fx_base','calc_plan','calc_actual','link_actual_to_plan',
    'is_approved','norm_status','approved_mask','match_actual_to_plan','link_actual_to_plan','purchase_invoice_to_actual','work_group_summary','work_item_variance','ifc_qto_takeoff_candidates','dxf_layer_takeoff_candidates','find_duplicate_dxf_layers','read_dxf','entity_measure'
}
WANTED_ASSIGNMENTS = {'PLAN','ACTUAL','WORK_GROUPS','WORK_RULES','UNIT','CURRENCIES','_SCOPE_LOCKED_SUBCATEGORIES'}
NAMESPACE = {'pd':pd,'np':np,'math':math,'re':re,'difflib':difflib,'uuid':uuid,'datetime':datetime}
selected=[]
for node in TREE.body:
    if isinstance(node, ast.FunctionDef) and node.name in WANTED_FUNCS:
        selected.append(node)
    elif isinstance(node, (ast.Assign, ast.AnnAssign)):
        targets=[]
        if isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name): targets.append(t.id)
        elif isinstance(node.target, ast.Name): targets.append(node.target.id)
        if any(t in WANTED_ASSIGNMENTS for t in targets): selected.append(node)
MODULE=ast.Module(body=selected,type_ignores=[])
exec(compile(MODULE,'<v15-pure-logic>','exec'),NAMESPACE)
NAMESPACE['now_tr']=lambda: datetime(2026,11,14)

class V15LogicTests(unittest.TestCase):
    def test_work_group_classification(self):
        classify=NAMESPACE['classify_work_item']
        self.assertEqual(classify('Grobeton'),('Kaba İşler','Grobeton'))
        self.assertEqual(classify('C30/37 temel betonu'),('Kaba İşler','Temel / Radye Betonu'))
        self.assertEqual(classify('B420C nervürlü inşaat demiri'),('Kaba İşler','Donatı / İnşaat Demiri'))
        self.assertEqual(classify('Kalıp ve demir işçiliği'),('Kaba İşler','Kalıp + Demir İşçiliği'))
        self.assertEqual(classify('Bitümlü temel membranı'),('Yalıtım İşleri','Temel / Perde Su Yalıtımı'))
        self.assertEqual(classify('Temel altı bitümlü membran 4 mm'),('Yalıtım İşleri','Temel / Perde Su Yalıtımı'))
        self.assertEqual(classify('Grobeton üstü bitümlü membran'),('Yalıtım İşleri','Temel / Perde Su Yalıtımı'))
        self.assertEqual(classify('Alüminyum kompozit cephe'),('Cephe İşleri','Cephe Kaplama'))
        self.assertEqual(classify('İç cephe boya'),('İnce İşler','Boya / Duvar Kaplama'))
        self.assertEqual(classify('Seramik cephe kaplama'),('Cephe İşleri','Cephe Kaplama'))
        self.assertEqual(classify('Çatı membran su yalıtımı'),('Yalıtım İşleri','Çatı Su Yalıtımı'))
        self.assertEqual(classify('Temiz su PPRC boru'),('Mekanik Tesisat','Temiz / Pis Su Tesisatı'))
        self.assertEqual(classify('Elektrik kablo'),('Elektrik Tesisat','Kuvvetli Akım'))

    def test_unit_conversion(self):
        conv=NAMESPACE['convert_qty']
        self.assertAlmostEqual(conv(5000,'kg','ton'),5.0)
        self.assertAlmostEqual(conv(5,'ton','kg'),5000.0)
        self.assertAlmostEqual(conv(100,'m²','m2'),100.0)
        self.assertIsNone(conv(5,'m3','m2'))

    def test_match_by_wbs_and_unit_conversion(self):
        PLAN=NAMESPACE['PLAN']
        plan=pd.DataFrame([{'_id':'p1','WBS':'KAB.01','Poz No':'','İş Kalemi':'İnşaat demiri','Birim':'ton','Miktar':18,'Birim Fiyat':32000,'Para Birimi':'TRY','Kur':1,'Metraj Durumu':'Onaylandı'}])
        got=NAMESPACE['match_actual_to_plan']({'WBS':'KAB.01','Poz No':'','İş Kalemi':'Demir alımı','Birim':'kg'},plan)
        self.assertEqual(got['status'],'KESİN WBS EŞLEŞMESİ')
        self.assertEqual(got['matches'][0]['_id'],'p1')

    def test_fuzzy_material_match_and_technical_rejection(self):
        plan=pd.DataFrame([{'_id':'p1','WBS':'KAB.01','Poz No':'','İş Kalemi':'B420C inşaat demiri donatı','Birim':'ton','Miktar':18,'Birim Fiyat':32000,'Para Birimi':'TRY','Kur':1,'Metraj Durumu':'Onaylandı'}])
        got=NAMESPACE['match_actual_to_plan']({'WBS':'','Poz No':'','İş Kalemi':'B420C donatı','Birim':'ton'},plan)
        self.assertTrue(got['matches'])
        self.assertGreaterEqual(got['score'],0.72)
        concrete=pd.DataFrame([{'_id':'p2','WBS':'KAB.02','Poz No':'','İş Kalemi':'C30 beton','Birim':'m3','Miktar':100,'Birim Fiyat':3000,'Para Birimi':'TRY','Kur':1,'Metraj Durumu':'Onaylandı'}])
        bad=NAMESPACE['match_actual_to_plan']({'WBS':'','Poz No':'','İş Kalemi':'C25 beton','Birim':'m3'},concrete)
        self.assertFalse(bad['matches'])

    def test_duplicate_wbs_never_auto_selects(self):
        plan=pd.DataFrame([
            {'_id':'p1','WBS':'KAB.01','Poz No':'','İş Kalemi':'Grobeton','Birim':'m3'},
            {'_id':'p2','WBS':'KAB.01','Poz No':'','İş Kalemi':'Grobeton','Birim':'m3'},
        ])
        got=NAMESPACE['match_actual_to_plan']({'WBS':'KAB.01','Poz No':'','İş Kalemi':'Grobeton','Birim':'m3'},plan)
        self.assertIn('TEKRARLI',got['status'])
        self.assertEqual(len(got['matches']),2)

    def test_ifc_qto_candidates_use_explicit_quantities_only(self):
        ifc=pd.DataFrame([
            {'GlobalId':'wall-guid','IFC Türü':'IfcWall','Ad':'Interior wall','ObjectType':'Masonry','NetSideArea':120.0},
            {'GlobalId':'rebar-guid','IFC Türü':'IfcReinforcingBar','Ad':'Rebar set','ObjectType':'B420C','NetWeightKg':2500.0},
            {'GlobalId':'pipe-guid','IFC Türü':'IfcPipeSegment','Ad':'Cold water pipe','ObjectType':'PPRC','Length':42.0},
            {'GlobalId':'bar-no-qto','IFC Türü':'IfcReinforcingBar','Ad':'No weight','ObjectType':'B420C','Length':100.0},
        ])
        got=NAMESPACE['ifc_qto_takeoff_candidates'](ifc,pd.DataFrame(),'TRY','test.ifc')
        self.assertEqual(len(got),3)
        self.assertIn(('Duvar yüzey alanı (IFC Qto)','m²'),set(zip(got['İş Kalemi'],got['Birim'])))
        self.assertIn(('Donatı ağırlığı (IFC Qto)','kg'),set(zip(got['İş Kalemi'],got['Birim'])))
        self.assertIn(('Boru tesisatı uzunluğu (IFC Qto Length)','m'),set(zip(got['İş Kalemi'],got['Birim'])))
        self.assertTrue(got['WBS'].fillna('').eq('').all())
        self.assertTrue(got['Metraj Durumu'].eq('Kontrol Bekliyor').all())
        self.assertTrue(got['Birim Fiyat'].isna().all())


    def test_ifc_qto_classifies_slab_concrete_and_roof_separately(self):
        data=pd.DataFrame([
            {'GlobalId':'slab-volume','IFC Türü':'IfcSlab','Ad':'Concrete slab','ObjectType':'C30/37','NetVolume':42.0,'NetArea':140.0},
            {'GlobalId':'roof','IFC Türü':'IfcRoof','Ad':'Roof covering','ObjectType':'Metal roof','NetArea':120.0},
        ])
        got=NAMESPACE['ifc_qto_takeoff_candidates'](data,pd.DataFrame(),'TRY','classification.ifc')
        slab_volume=got[got['Qto Alanı']=='NetVolume'].iloc[0]
        slab_area=got[got['Qto Alanı']=='NetArea'].iloc[0]
        roof=got[got['IFC GlobalId']=='roof'].iloc[0]
        self.assertEqual(slab_volume['Ana İş Grubu'],'Kaba İşler')
        self.assertEqual(slab_volume['Alt Kategori'],'Döşeme Betonu')
        self.assertEqual(slab_volume['Birim'],'m³')
        self.assertEqual(slab_area['Ana İş Grubu'],'Kaba İşler')
        self.assertIn('Kontrol',slab_area['Alt Kategori'])
        self.assertEqual(roof['Ana İş Grubu'],'Çatı ve Peyzaj')
        self.assertEqual(roof['Alt Kategori'],'Çatı İmalatları')

    def test_dxf_import_detects_same_file_layer_reimport(self):
        existing=pd.DataFrame([{'Açıklama':'DXF katmanı: A-MEMBRAN · Kaynak: a.dxf · DXF SHA256: abc123 · metraj'}])
        candidates=pd.DataFrame([{'Kaynak Layer':'A-MEMBRAN'},{'Kaynak Layer':'A-BORU'}])
        self.assertEqual(NAMESPACE['find_duplicate_dxf_layers'](existing,candidates,'abc123'),['A-MEMBRAN'])
        self.assertEqual(NAMESPACE['find_duplicate_dxf_layers'](existing,candidates,'changed-hash'),[])

    def test_ifc_qto_candidates_aggregate_by_wbs_and_prevent_collisions(self):
        rows=pd.DataFrame([
            {'_id':'a','WBS':'KAB.01.01','Poz No':'','İş Kalemi':'Döşeme hacmi','Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler','Alt Kategori':'Betonarme','Birim':'m³','Miktar':12.5,'Birim Fiyat':3000,'Para Birimi':'TRY','Fiyat Yılı':2026,'Fiyat Kaynağı':'Kaynak A','Açıklama':'IFC GlobalId: wall-1 · Eleman: Slab · Qto alanı: NetVolume.','IFC GlobalId':'wall-1'},
            {'_id':'b','WBS':'KAB.01.01','Poz No':'','İş Kalemi':'Döşeme hacmi','Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler','Alt Kategori':'Betonarme','Birim':'m³','Miktar':7.5,'Birim Fiyat':3000,'Para Birimi':'TRY','Fiyat Yılı':2026,'Fiyat Kaynağı':'Kaynak A','Açıklama':'IFC GlobalId: wall-2 · Eleman: Slab · Qto alanı: NetVolume.','IFC GlobalId':'wall-2'},
        ])
        agg=NAMESPACE['aggregate_ifc_qto_candidates'](rows,pd.DataFrame())
        self.assertEqual(len(agg),1)
        self.assertAlmostEqual(agg.iloc[0]['Miktar'],20.0)
        self.assertIn('wall-1',agg.iloc[0]['Açıklama'])
        self.assertIn('wall-2',agg.iloc[0]['Açıklama'])
        with self.assertRaises(ValueError):
            NAMESPACE['aggregate_ifc_qto_candidates'](rows,pd.DataFrame([{'_id':'x','WBS':'KAB.01.01'}]))
        conflict=rows.copy(); conflict.loc[1,'İş Kalemi']='Kolon betonu hacmi'
        with self.assertRaises(ValueError):
            NAMESPACE['aggregate_ifc_qto_candidates'](conflict,pd.DataFrame())
        conflict=rows.copy(); conflict.loc[1,'Poz No']='15.150.1001'
        with self.assertRaises(ValueError):
            NAMESPACE['aggregate_ifc_qto_candidates'](conflict,pd.DataFrame())

    def test_real_dxf_geometry_extracts_scaled_length_and_closed_area(self):
        import ezdxf
        with tempfile.TemporaryDirectory() as td:
            path=Path(td)/'sample.dxf'
            doc=ezdxf.new('R2010')
            doc.header['$INSUNITS']=6  # metres
            msp=doc.modelspace()
            msp.add_line((0,0),(3,4),dxfattribs={'layer':'M-PPRC-BORU'})
            msp.add_lwpolyline([(0,0),(10,0),(10,5),(0,5)],close=True,dxfattribs={'layer':'A-MEMBRAN'})
            msp.add_lwpolyline([(20,0,1),(30,0,0)],format='xyb',close=True,dxfattribs={'layer':'A-BULGE'})
            hatch=msp.add_hatch(dxfattribs={'layer':'A-HATCH-MEMBRAN'})
            hatch.paths.add_polyline_path([(0,10),(4,10),(4,15),(0,15)],is_closed=True)
            doc.saveas(path)
            result=NAMESPACE['read_dxf'](path)
            self.assertEqual(result['unit_label'],'m')
            detail=pd.DataFrame(result['detail'])
            pipe=detail[detail['Katman']=='M-PPRC-BORU'].iloc[0]
            membrane=detail[detail['Katman']=='A-MEMBRAN'].iloc[0]
            bulge=detail[detail['Katman']=='A-BULGE'].iloc[0]
            hatch=detail[detail['Katman']=='A-HATCH-MEMBRAN'].iloc[0]
            self.assertAlmostEqual(pipe['Uzunluk (m)'],5.0)
            self.assertAlmostEqual(membrane['Alan (m²)'],50.0)
            self.assertAlmostEqual(hatch['Alan (m²)'],20.0)
            self.assertGreater(bulge['Uzunluk (m)'],25.5)
            self.assertLess(bulge['Uzunluk (m)'],25.9)
            self.assertGreater(bulge['Alan (m²)'],39.0)
            self.assertLess(bulge['Alan (m²)'],39.5)
            candidates=NAMESPACE['dxf_layer_takeoff_candidates'](result,pd.DataFrame(),'TRY','sample.dxf','sha-test')
            self.assertEqual(set(candidates['Birim']),{'m','m²'})
            self.assertEqual(len(candidates),3)
            self.assertTrue(candidates['Açıklama'].str.contains('DXF SHA256: sha-test').all())

    def test_dxf_layer_candidates_prefer_hatch_and_do_not_infer_concrete_volume(self):
        detail=pd.DataFrame([
            {'Tip':'HATCH','Katman':'A-MEMBRAN','Uzunluk (m)':None,'Alan (m²)':100.0},
            {'Tip':'LWPOLYLINE','Katman':'A-MEMBRAN','Uzunluk (m)':30.0,'Alan (m²)':100.0},
            {'Tip':'LINE','Katman':'M-PPRC-BORU','Uzunluk (m)':50.0,'Alan (m²)':None},
            {'Tip':'HATCH','Katman':'S-BETON','Uzunluk (m)':None,'Alan (m²)':75.0},
        ])
        result={'detail':detail.to_dict('records'),'factor':1.0}
        got=NAMESPACE['dxf_layer_takeoff_candidates'](result,pd.DataFrame(),'TRY','test.dxf')
        self.assertEqual(len(got),2)
        membrane=got[got['Kaynak Layer']=='A-MEMBRAN'].iloc[0]
        pipe=got[got['Kaynak Layer']=='M-PPRC-BORU'].iloc[0]
        self.assertAlmostEqual(membrane['Miktar'],100.0)
        self.assertEqual(membrane['Birim'],'m²')
        self.assertAlmostEqual(pipe['Miktar'],50.0)
        self.assertEqual(pipe['Birim'],'m')
        self.assertTrue((got['Metraj Durumu']=='Kontrol Bekliyor').all())
        self.assertTrue(got['Birim Fiyat'].isna().all())

    def test_link_actual_to_plan_preserves_total_and_converts_kg_to_ton(self):
        actual={'WBS':'','İş Kalemi':'B420C nervürlü inşaat demiri','Birim':'kg','Gerçekleşen Miktar':5000,
                'Gerçekleşen Birim Fiyat':34.5,'Ana Kategori':'Betonarme İmalatlar','Açıklama':'İrsaliye 123'}
        target={'_id':'p1','WBS':'KAB.01','İş Kalemi':'B420C inşaat demiri donatı','Birim':'ton',
                'Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler','Alt Kategori':'Donatı / İnşaat Demiri'}
        mapped=NAMESPACE['link_actual_to_plan'](actual,target)
        self.assertAlmostEqual(mapped['Gerçekleşen Miktar'],5.0)
        self.assertAlmostEqual(mapped['Gerçekleşen Birim Fiyat'],34500.0)
        self.assertAlmostEqual(mapped['Gerçekleşen Miktar']*mapped['Gerçekleşen Birim Fiyat'],172500.0)
        self.assertEqual(mapped['WBS'],'KAB.01')
        self.assertEqual(mapped['Ana İş Grubu'],'Kaba İşler')

    def test_link_actual_to_plan_requires_explicit_technical_override(self):
        actual={'İş Kalemi':'C25 temel betonu','Birim':'m³','Gerçekleşen Miktar':10,'Gerçekleşen Birim Fiyat':2500}
        target={'WBS':'KAB.02','İş Kalemi':'C30/37 temel betonu','Birim':'m³','Ana İş Grubu':'Kaba İşler'}
        with self.assertRaises(ValueError): NAMESPACE['link_actual_to_plan'](actual,target)
        mapped=NAMESPACE['link_actual_to_plan'](actual,target,allow_technical_override=True)
        self.assertEqual(mapped['WBS'],'KAB.02')
        self.assertIn('teyit edildi',mapped['Açıklama'])

    def test_exact_wbs_does_not_bypass_technical_compatibility(self):
        plan=pd.DataFrame([{'_id':'p1','WBS':'KAB.02','Poz No':'','İş Kalemi':'C30/37 temel betonu','Birim':'m³'}])
        got=NAMESPACE['match_actual_to_plan']({'WBS':'KAB.02','Poz No':'','İş Kalemi':'C25 temel betonu','Birim':'m³'},plan)
        self.assertIn('TEKNİK ÖZELLİK UYUŞMAZLIĞI',got['status'])
        self.assertFalse(got['matches'])

    def test_item_variance_reconciles_budget_units_and_flags_unlinked_actual(self):
        plan=pd.DataFrame([{'_id':'p1','WBS':'KAB.01','İş Kalemi':'B420C inşaat demiri','Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler','Birim':'ton','Miktar':18,'Birim Fiyat':32000,'Para Birimi':'TRY','Kur':1,'Metraj Durumu':'Onaylandı'}])
        actual=pd.DataFrame([
            {'_id':'a1','WBS':'KAB.01','İş Kalemi':'B420C nervürlü demir alımı','Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler','Birim':'kg','Gerçekleşen Miktar':5000,'Gerçekleşen Birim Fiyat':34.5,'Para Birimi':'TRY','Kur':1,'Durum':'Onaylandı'},
            {'_id':'a2','WBS':'UNKNOWN','İş Kalemi':'Nakliye','Birim':'sefer','Gerçekleşen Miktar':1,'Gerçekleşen Birim Fiyat':1000,'Para Birimi':'TRY','Kur':1,'Durum':'Onaylandı'},
            {'_id':'a3','WBS':'','İş Kalemi':'Nakliye','Birim':'sefer','Gerçekleşen Miktar':1,'Gerçekleşen Birim Fiyat':500,'Para Birimi':'TRY','Kur':1,'Durum':'Onaylandı'}
        ])
        result=NAMESPACE['work_item_variance'](plan,actual,'TRY')
        linked=result[result['WBS']=='KAB.01'].iloc[0]
        self.assertAlmostEqual(linked['Gerçekleşen Miktar'],5.0)
        self.assertAlmostEqual(linked['Maliyet Sapması'],172500-576000)
        orphan=result[result['WBS']=='UNKNOWN'].iloc[0]
        self.assertEqual(orphan['Teknik Kontrol'],'WBS EŞLEŞMESİ YOK')
        no_wbs=result[result['WBS']=='WBS YOK'].iloc[0]
        self.assertEqual(no_wbs['Teknik Kontrol'],'WBS EŞLEŞMESİ YOK')

    def test_different_concrete_work_scopes_never_auto_match_even_with_same_grade(self):
        plan=pd.DataFrame([{'_id':'p1','WBS':'KAB.02','Poz No':'','İş Kalemi':'C30/37 kolon betonu','Birim':'m³'}])
        actual={'WBS':'KAB.02','Poz No':'','İş Kalemi':'C30/37 temel betonu','Birim':'m³'}
        got=NAMESPACE['match_actual_to_plan'](actual,plan)
        self.assertIn('İMALAT KAPSAMI UYUŞMAZLIĞI',got['status'])
        self.assertFalse(got['matches'])
        actual['WBS']=''
        fuzzy=NAMESPACE['match_actual_to_plan'](actual,plan)
        self.assertFalse(fuzzy['matches'])

    def test_scope_compatibility_detects_distinct_electrical_systems(self):
        compatible=NAMESPACE['work_scope_compatible']
        self.assertFalse(compatible('Zayıf akım data kablosu','Kuvvetli akım enerji kablosu'))
        self.assertFalse(compatible('Temel membran su yalıtımı','Teras sürme su yalıtımı'))
        self.assertTrue(compatible('B420C nervürlü demir','B420C inşaat demiri donatı'))

    def test_price_match_requires_compatible_scope_unit_currency_and_technical_grade(self):
        row={'Poz No':'15.150.1001','İş Kalemi':'C30/37 temel betonu','Birim':'m³','Para Birimi':'TRY','Ana Kategori':'Betonarme İmalatlar'}
        library=pd.DataFrame([
            {'Poz No':'15.150.1001','İş Kalemi':'C30/37 temel betonu','Birim':'m3','Para Birimi':'TRY','Birim Fiyat':3500,'Fiyat Yılı':2026,'Kaynak':'YFK 2026','Aktif':True},
            {'Poz No':'15.150.1002','İş Kalemi':'C30/37 kolon betonu','Birim':'m3','Para Birimi':'TRY','Birim Fiyat':3600,'Fiyat Yılı':2026,'Kaynak':'YFK 2026','Aktif':True},
        ])
        hit=NAMESPACE['price_match'](row,library)
        self.assertEqual(hit['status'],'YÜKSEK EŞLEŞME — ONAY GEREKLİ')
        self.assertEqual(hit['price'],3500)
        bad_unit=library.iloc[[0]].copy(); bad_unit.loc[bad_unit.index[0],'Birim']='ton'
        self.assertIsNone(NAMESPACE['price_match'](row,bad_unit).get('price'))
        bad_currency=library.iloc[[0]].copy(); bad_currency.loc[bad_currency.index[0],'Para Birimi']='EUR'
        self.assertIsNone(NAMESPACE['price_match'](row,bad_currency).get('price'))
        bad_scope=library.iloc[[1]].copy(); bad_scope.loc[bad_scope.index[0],'Poz No']='15.150.1001'
        self.assertIsNone(NAMESPACE['price_match'](row,bad_scope).get('price'))
        bad_code=library.iloc[[0]].copy(); bad_code.loc[bad_code.index[0],'Poz No']='15.150.9999'
        self.assertIsNone(NAMESPACE['price_match'](row,bad_code).get('price'))

    def test_price_match_marks_old_price_and_rejects_future_validity(self):
        row={'Poz No':'','İş Kalemi':'B420C inşaat demiri donatı','Birim':'ton','Para Birimi':'TRY','Ana Kategori':'Betonarme İmalatlar'}
        stale=pd.DataFrame([{'Poz No':'','İş Kalemi':'B420C nervürlü inşaat demiri donatı','Birim':'ton','Para Birimi':'TRY','Birim Fiyat':32000,'Fiyat Yılı':2024,'Kaynak':'Eski teklif','Aktif':True}])
        self.assertIn('ESKİ FİYAT',NAMESPACE['price_match'](row,stale)['status'])
        future=stale.copy(); future.loc[0,'Fiyat Yılı']=2026; future.loc[0,'Geçerlilik Başlangıcı']='2027-01-01'
        self.assertIsNone(NAMESPACE['price_match'](row,future).get('price'))

    def test_explicit_work_group_overrides_legacy_mixed_category(self):
        compatible=NAMESPACE['work_scope_compatible']
        self.assertFalse(compatible('Boru tesisatı uzunluğu','Elektrik kablo tesisatı',
                                    'Mekanik ve Elektrik','Mekanik ve Elektrik',
                                    'Mekanik Tesisat','Elektrik Tesisat','Boru Tesisatı','Kablo Tesisatı'))
        self.assertTrue(compatible('PPRC boru','Boru tesisatı Qto',
                                   'Mekanik ve Elektrik','Mekanik ve Elektrik',
                                   'Mekanik Tesisat','Mekanik Tesisat','Temiz / Pis Su Tesisatı','Temiz / Pis Su Tesisatı'))

    def test_technical_compatibility_detects_diameter_and_finish_differences(self):
        compatible=NAMESPACE['tech_compatible']
        self.assertFalse(compatible('B420C Ø10 nervürlü demir','B420C Ø12 nervürlü demir'))
        self.assertFalse(compatible('A2 paslanmaz vida M10x90','A4 paslanmaz vida M10x90'))
        self.assertFalse(compatible('Galvanizli çelik profil 3 mm','Çelik profil 3 mm'))
        self.assertTrue(compatible('C30/37 temel betonu','C30/37 temel betonu'))

    def test_takeoff_updates_existing_wbs_as_pending_revision_and_preserves_price(self):
        plan=pd.DataFrame([{'_id':'p1','WBS':'KAB.01','Poz No':'','İş Kalemi':'Grobeton','Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler','Alt Kategori':'Grobeton','Birim':'m³','Miktar':18,'Birim Fiyat':2500,'Para Birimi':'TRY','Kur':1,'Fiyat Yılı':2026,'Fiyat Kaynağı':'Onaylı fiyat','Metraj Durumu':'Onaylandı','Açıklama':'İlk metraj'}])
        candidates=pd.DataFrame([{'_id':'c1','WBS':'KAB.01','İş Kalemi':'Grobeton','Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler','Alt Kategori':'Grobeton','Birim':'m³','Miktar':21.5,'Birim Fiyat':np.nan,'Para Birimi':'TRY','Kur':1,'Metraj Durumu':'Kontrol Bekliyor','Açıklama':'DXF layer A-GROBETON'}])
        merged,added,updated=NAMESPACE['merge_takeoff_candidates'](plan,candidates,'DXF')
        self.assertEqual((added,updated),(0,1))
        self.assertEqual(len(merged),1)
        self.assertAlmostEqual(merged.iloc[0]['Miktar'],21.5)
        self.assertEqual(merged.iloc[0]['Birim Fiyat'],2500)
        self.assertEqual(merged.iloc[0]['Metraj Durumu'],'Kontrol Bekliyor')
        self.assertIn('önceki miktar 18',merged.iloc[0]['Açıklama'])

    def test_takeoff_adds_new_wbs_and_converts_units_to_existing_budget_unit(self):
        plan=pd.DataFrame([{'_id':'p1','WBS':'KAB.10','İş Kalemi':'B420C inşaat demiri donatı','Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler','Alt Kategori':'Donatı / İnşaat Demiri','Birim':'ton','Miktar':18,'Birim Fiyat':32000,'Para Birimi':'TRY','Kur':1,'Metraj Durumu':'Onaylandı'}])
        candidates=pd.DataFrame([{'_id':'c1','WBS':'KAB.10','İş Kalemi':'B420C nervürlü inşaat demiri','Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler','Alt Kategori':'Donatı / İnşaat Demiri','Birim':'kg','Miktar':5000,'Birim Fiyat':np.nan,'Para Birimi':'TRY'},
                                 {'_id':'c2','WBS':'DXF.NEW01','İş Kalemi':'Membran / Su Yalıtımı','Ana Kategori':'Yalıtım / Çatı / Cephe','Ana İş Grubu':'Yalıtım İşleri','Alt Kategori':'Temel / Perde Su Yalıtımı','Birim':'m²','Miktar':110,'Birim Fiyat':np.nan,'Para Birimi':'TRY'}])
        merged,added,updated=NAMESPACE['merge_takeoff_candidates'](plan,candidates,'IFC')
        self.assertEqual((added,updated),(1,1))
        linked=merged[merged['WBS']=='KAB.10'].iloc[0]
        self.assertAlmostEqual(linked['Miktar'],5.0)
        self.assertEqual(linked['Birim'],'ton')
        self.assertEqual(linked['Metraj Durumu'],'Kontrol Bekliyor')
        self.assertEqual(len(merged),2)
        self.assertEqual(merged[merged['WBS']=='DXF.NEW01'].iloc[0]['Metraj Durumu'],'Kontrol Bekliyor')

    def test_takeoff_does_not_overwrite_existing_wbs_with_wrong_scope_or_duplicates(self):
        base=pd.DataFrame([{'_id':'p1','WBS':'KAB.02','İş Kalemi':'C30/37 kolon betonu','Birim':'m³','Miktar':100,'Birim Fiyat':3000,'Metraj Durumu':'Onaylandı'}])
        candidate=pd.DataFrame([{'_id':'c1','WBS':'KAB.02','İş Kalemi':'C30/37 temel betonu','Birim':'m³','Miktar':20}])
        with self.assertRaises(ValueError): NAMESPACE['merge_takeoff_candidates'](base,candidate,'DXF')
        duplicate=pd.concat([base,base.assign(_id='p2')],ignore_index=True)
        valid=pd.DataFrame([{'_id':'c2','WBS':'KAB.02','İş Kalemi':'C30/37 kolon betonu','Birim':'m³','Miktar':20}])
        with self.assertRaises(ValueError): NAMESPACE['merge_takeoff_candidates'](duplicate,valid,'DXF')

    def test_scope_override_must_be_explicit_for_manual_actual_link(self):
        actual={'İş Kalemi':'C30/37 temel betonu','Birim':'m³','Gerçekleşen Miktar':10,'Gerçekleşen Birim Fiyat':3000}
        target={'WBS':'KAB.02','İş Kalemi':'C30/37 kolon betonu','Birim':'m³','Ana İş Grubu':'Kaba İşler'}
        with self.assertRaises(ValueError): NAMESPACE['link_actual_to_plan'](actual,target)
        linked=NAMESPACE['link_actual_to_plan'](actual,target,allow_scope_override=True)
        self.assertEqual(linked['WBS'],'KAB.02')
        self.assertIn('kullanıcı tarafından teyit edildi',linked['Açıklama'])

    def test_reconciliation_excludes_cost_when_scope_is_wrong(self):
        plan=pd.DataFrame([{'_id':'p1','WBS':'KAB.02','İş Kalemi':'C30/37 kolon betonu','Ana Kategori':'Betonarme İmalatlar','Birim':'m³','Miktar':100,'Birim Fiyat':3000,'Para Birimi':'TRY','Kur':1,'Metraj Durumu':'Onaylandı'}])
        actual=pd.DataFrame([{'_id':'a1','WBS':'KAB.02','İş Kalemi':'C30/37 temel betonu','Ana Kategori':'Betonarme İmalatlar','Birim':'m³','Gerçekleşen Miktar':10,'Gerçekleşen Birim Fiyat':3000,'Para Birimi':'TRY','Kur':1,'Durum':'Onaylandı'}])
        result=NAMESPACE['work_item_variance'](plan,actual,'TRY').iloc[0]
        self.assertEqual(result['Teknik Kontrol'],'TEKNİK UYUŞMAZLIK')
        self.assertTrue(pd.isna(result['Gerçekleşen Tutar']))
        self.assertTrue(pd.isna(result['Maliyet Sapması']))

    def test_takeoff_same_wbs_cannot_mix_different_imperial_items(self):
        candidates=pd.DataFrame([{'_id':'c1','WBS':'KAB.01','İş Kalemi':'Grobeton','Birim':'m³','Miktar':10},
                                 {'_id':'c2','WBS':'KAB.01','İş Kalemi':'Temel betonu','Birim':'m³','Miktar':10}])
        with self.assertRaises(ValueError): NAMESPACE['merge_takeoff_candidates'](pd.DataFrame(),candidates,'DXF')

    def test_duplicate_budget_wbs_is_not_double_counted(self):
        plan=pd.DataFrame([
            {'_id':'p1','WBS':'KAB.01','İş Kalemi':'Grobeton','Birim':'m3','Miktar':10,'Birim Fiyat':1000,'Para Birimi':'TRY','Kur':1,'Metraj Durumu':'Onaylandı'},
            {'_id':'p2','WBS':'KAB.01','İş Kalemi':'Grobeton','Birim':'m3','Miktar':20,'Birim Fiyat':1000,'Para Birimi':'TRY','Kur':1,'Metraj Durumu':'Onaylandı'}
        ])
        actual=pd.DataFrame([{'_id':'a1','WBS':'KAB.01','İş Kalemi':'Grobeton','Birim':'m3','Gerçekleşen Miktar':5,'Gerçekleşen Birim Fiyat':1100,'Para Birimi':'TRY','Kur':1,'Durum':'Onaylandı'}])
        result=NAMESPACE['work_item_variance'](plan,actual,'TRY')
        self.assertEqual(len(result),2)
        self.assertTrue(result['Teknik Kontrol'].eq('WBS TEKRARLI — MUTABAKAT YAPILAMAZ').all())
        self.assertTrue(result['Gerçekleşen Tutar'].isna().all())

    def test_approved_group_budget_vs_actual(self):
        plan=pd.DataFrame([{'_id':'p1','WBS':'KAB.01','İş Kalemi':'Grobeton','Ana Kategori':'Betonarme İmalatlar','Birim':'m3','Miktar':20,'Birim Fiyat':2500,'Para Birimi':'TRY','Kur':1,'Metraj Durumu':'Onaylandı'}])
        actual=pd.DataFrame([{'_id':'a1','WBS':'KAB.01','İş Kalemi':'Grobeton alımı','Ana Kategori':'Betonarme İmalatlar','Birim':'m3','Gerçekleşen Miktar':18,'Gerçekleşen Birim Fiyat':2700,'Para Birimi':'TRY','Kur':1,'Tarih':'2026-11-14','Durum':'Onaylandı'}])
        summary=NAMESPACE['work_group_summary'](plan,actual,'TRY')
        row=summary[summary['Ana İş Grubu']=='Kaba İşler'].iloc[0]
        self.assertAlmostEqual(row['Onaylı Bütçe'],50000)
        self.assertAlmostEqual(row['Onaylı Gerçekleşen'],48600)
        self.assertAlmostEqual(row['Sapma'],-1400)

    def test_generic_membrane_is_not_auto_assigned_to_foundation(self):
        classify=NAMESPACE['classify_work_item']
        self.assertEqual(classify('Membran / Su Yalıtımı'),('Yalıtım İşleri','Uygulama Yeri Belirsiz — Yalıtım'))
        self.assertEqual(classify('Bitümlü temel membranı'),('Yalıtım İşleri','Temel / Perde Su Yalıtımı'))
        self.assertFalse(NAMESPACE['work_scope_compatible']('Membran / Su Yalıtımı','Bitümlü temel membranı'))

    def test_purchase_invoice_creates_quantity_based_pending_actual(self):
        plan=pd.DataFrame([{
            '_id':'p1','WBS':'KAB.01','Poz No':'','İş Kalemi':'B420C inşaat demiri donatı',
            'Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler',
            'Alt Kategori':'Donatı / İnşaat Demiri','Birim':'ton','Miktar':18,
            'Birim Fiyat':32000,'Para Birimi':'TRY','Kur':1,'Metraj Durumu':'Onaylandı'
        }])
        purchase={
            'Sipariş No':'PO-001','WBS':'KAB.01','Poz No':'','Malzeme / Hizmet':'B420C nervürlü inşaat demiri',
            'Tedarikçi':'Çelik Ltd.','Birim':'kg','Para Birimi':'TRY','Kur':1,
            'Faturalanan Miktar':5000,'Faturalanan Tutar':172500,
            'Fatura No':'FAT-001','Fatura Tarihi':'14.11.2026'
        }
        actual=NAMESPACE['purchase_invoice_to_actual'](purchase,plan,'TRY')
        self.assertEqual(actual['WBS'],'KAB.01')
        self.assertEqual(actual['Birim'],'ton')
        self.assertAlmostEqual(actual['Gerçekleşen Miktar'],5.0)
        self.assertAlmostEqual(actual['Gerçekleşen Birim Fiyat'],34500.0)
        self.assertAlmostEqual(actual['Gerçekleşen Tutar'],172500.0)
        self.assertEqual(actual['Durum'],'Kontrol Bekliyor')
        self.assertEqual(actual['Belge No'],'FAT-001')
        actual_frame=NAMESPACE['calc_actual'](pd.DataFrame([actual]),'TRY')
        self.assertAlmostEqual(actual_frame.iloc[0]['Baz Para Birimi Tutarı'],172500.0)

    def test_purchase_invoice_applies_user_supplied_fx_without_changing_foreign_amount(self):
        plan=pd.DataFrame([{
            '_id':'p1','WBS':'KAB.01','İş Kalemi':'B420C inşaat demiri donatı',
            'Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler',
            'Alt Kategori':'Donatı / İnşaat Demiri','Birim':'ton'
        }])
        purchase={'WBS':'KAB.01','Malzeme / Hizmet':'B420C inşaat demiri','Birim':'kg','Para Birimi':'EUR',
                  'Kur':40.0,'Kur Kaynağı':'Banka dekontu','Faturalanan Miktar':1000,
                  'Faturalanan Tutar':1000,'Fatura No':'FAT-EUR','Fatura Tarihi':'14.11.2026'}
        actual=NAMESPACE['purchase_invoice_to_actual'](purchase,plan,'TRY')
        self.assertEqual(actual['Para Birimi'],'EUR')
        self.assertAlmostEqual(actual['Gerçekleşen Tutar'],1000.0)
        self.assertAlmostEqual(actual['Kur'],40.0)
        actual_frame=NAMESPACE['calc_actual'](pd.DataFrame([actual]),'TRY')
        self.assertAlmostEqual(actual_frame.iloc[0]['Baz Para Birimi Tutarı'],40000.0)

    def test_purchase_invoice_rejects_missing_quantity_duplicate_wbs_and_invalid_fx(self):
        plan=pd.DataFrame([{
            '_id':'p1','WBS':'KAB.01','İş Kalemi':'B420C inşaat demiri donatı',
            'Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler',
            'Alt Kategori':'Donatı / İnşaat Demiri','Birim':'ton'
        }])
        purchase={'WBS':'KAB.01','Malzeme / Hizmet':'B420C nervürlü inşaat demiri','Birim':'kg',
                  'Para Birimi':'TRY','Faturalanan Miktar':np.nan,'Faturalanan Tutar':172500,
                  'Fatura No':'FAT-002','Fatura Tarihi':'14.11.2026'}
        with self.assertRaisesRegex(ValueError,'Faturalanan miktarı'):
            NAMESPACE['purchase_invoice_to_actual'](purchase,plan,'TRY')
        duplicate=pd.concat([plan,plan.assign(_id='p2')],ignore_index=True)
        purchase['Faturalanan Miktar']=5000
        with self.assertRaisesRegex(ValueError,'tekil olarak'):
            NAMESPACE['purchase_invoice_to_actual'](purchase,duplicate,'TRY')
        purchase['WBS']='KAB.01'; purchase['Para Birimi']='EUR'; purchase['Kur']=0
        with self.assertRaisesRegex(ValueError,'geçerli işlem kuru'):
            NAMESPACE['purchase_invoice_to_actual'](purchase,plan,'TRY')

    def test_purchase_invoice_scope_and_technical_exceptions_require_explicit_confirmation(self):
        membrane_plan=pd.DataFrame([{
            '_id':'p1','WBS':'YAL.01','İş Kalemi':'Bitümlü temel membranı',
            'Ana Kategori':'Yalıtım / Çatı / Cephe','Ana İş Grubu':'Yalıtım İşleri',
            'Alt Kategori':'Temel / Perde Su Yalıtımı','Birim':'m²'
        }])
        purchase={'WBS':'YAL.01','Malzeme / Hizmet':'Membran','Birim':'m²','Para Birimi':'TRY',
                  'Faturalanan Miktar':100,'Faturalanan Tutar':25000,'Fatura No':'FAT-M',
                  'Fatura Tarihi':'14.11.2026'}
        with self.assertRaisesRegex(ValueError,'İmalat kapsamı uyuşmazlığı'):
            NAMESPACE['purchase_invoice_to_actual'](purchase,membrane_plan,'TRY')
        accepted=NAMESPACE['purchase_invoice_to_actual'](purchase,membrane_plan,'TRY',allow_scope_override=True)
        self.assertIn('kullanıcı tarafından teyit edildi',accepted['Açıklama'])

        concrete_plan=pd.DataFrame([{
            '_id':'p2','WBS':'KAB.02','İş Kalemi':'C30/37 temel betonu',
            'Ana Kategori':'Betonarme İmalatlar','Ana İş Grubu':'Kaba İşler',
            'Alt Kategori':'Temel / Radye Betonu','Birim':'m³'
        }])
        concrete_purchase={'WBS':'KAB.02','Malzeme / Hizmet':'C25 temel betonu','Birim':'m³','Para Birimi':'TRY',
                           'Faturalanan Miktar':10,'Faturalanan Tutar':25000,'Fatura No':'FAT-C',
                           'Fatura Tarihi':'14.11.2026'}
        with self.assertRaisesRegex(ValueError,'Teknik özellik uyuşmazlığı'):
            NAMESPACE['purchase_invoice_to_actual'](concrete_purchase,concrete_plan,'TRY')
        accepted=NAMESPACE['purchase_invoice_to_actual'](concrete_purchase,concrete_plan,'TRY',allow_technical_override=True)
        self.assertIn('Teknik özellik farkı kullanıcı tarafından teyit edildi',accepted['Açıklama'])

if __name__=='__main__':
    unittest.main(verbosity=2)
