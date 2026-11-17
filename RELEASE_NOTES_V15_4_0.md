# V15.4.0 — PDF Metraj + Web Piyasa Fiyat Araştırması + CAD Katman Eşleştirme

## V15.4.0
- PDF pafta dosyasından AI destekli, sayfa kanıtı/formül/güven skoru içeren metraj adayları üretme akışı eklendi.
- 1–24 sayfa analizi, 1–4 sayfalı AI API grupları; model quantity output strict schema validation.
- Dimension calculation formülleri tekrar hesaplanır; quantity/formula mismatch miktarı boşaltır.
- Source-backed Tavily Search API ile yerel tedarikçi/web fiyat kaynakları arama ve OpenAI ile açık fiyat ifadesi çıkarma eklendi.
- Modelin fiyat ifadesi kaynak ham metninde yoksa fiyat hücresi boş kalır; yalnızca source link saved.
- Teknik uyum onayı zorunlu internet price quote → company price library workflow.
- Yerel fiyat önerisini açıkça kabul etme ve sadece kaynak/yıl geçerli, eski olmayan öneriyi bütçe taslağına taşıma eklendi.
- DXF/DWG analizine manuel katman→imalat→ölçüm türü→WBS eşleştirme eklendi; HATCH/closed polyline area and polyline length logic.
- Miktar uydurmama, birim bilinmiyor veya CAD INSUNITS yoksa ölçmeme kontrolleri korundu.
- Tavily key configuration added to .env, Streamlit Secrets example and preflight warnings.
- Regression tests for manual CAD mapping, formula verification, web quote evidence, and explicit local price acceptance.

## Limits clearly retained
- IFC reads existing Qto quantities; no general-purpose solid geometry takeoff without Qto.
- DWG requires separately configured ODA worker.
- PDF AI takeoff is review-only and 24 page maximum per extraction run.
- Real OPENAI/TAVILY calls, production Streamlit, PostgreSQL, ClamAV, ODA and IFC integration were not tested in this offline environment.
