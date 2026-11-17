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

## Preserved V15.3.0
Company portfolio and purchase commitment/invoice/actual matching remain from the previous release.


## 15.5.0 — 2026-11-16
- Multi-provider AI gateway: OpenAI Responses, Anthropic Messages, Google Gemini Interactions, OpenAI-compatible endpoints.
- Opt-in cross-provider expert review with explicit data sharing consent; automatic fallback disabled by default.
- Expanded project review context and deterministic read-only controls across commercial, finance, site, material, labor and progress workflows.
- Gemini API key moved to x-goog-api-key header; provider request response storage defaults off.
- 90 pytest, 34 UI shell pages, 38 synthetic E2E passed; 2 external integrations skipped.
