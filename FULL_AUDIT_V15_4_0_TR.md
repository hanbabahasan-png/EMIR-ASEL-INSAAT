# V15.4.1 Uzman Teknik Denetim / Eksik Kapsam Tablosu

## Amaç
Bu sürümün odak noktası, kullanıcının asıl hedefi olan proje dokümanından metraj adayı oluşturma ve bu adayları izlenebilir birim fiyat kaynaklarına bağlamaktır.

## Tamamlanan kod akışları

1. **PDF AI takeoff:** PDF sayfa sayımı/şifre kontrolü, izinli sayfa seçimi, batch halinde text layer + page image gönderimi, strict JSON parse, page allow-list, recognised unit, evidence, positive quantity, measurement basis ve confidence validation.
2. **Formula check:** dimension_calculation satırlarında yalnızca + - * / parantez ve sayısal değerlerden oluşan güvenli AST evaluator çalışır; formula result AI quantity ile tolerans içinde uyuşmazsa miktar boşaltılır.
3. **Candidate store:** source filename/SHA/page/discipline/location/technical description/evidence/formula/confidence/WBS suggestion are stored in project-scoped `takeoff_candidates`.
4. **Budget transfer:** confidence >=0.60, positive amount, known unit, evidence and explicit user selection required. Transfer remains pending. Price stays empty unless user explicitly accepts a current, sourced local-library match. Baseline approval remains separate.
5. **Web price research:** Tavily advanced search with Turkey region, allow-listed HTTP/HTTPS result URLs, source titles/content/snippets and dates. OpenAI extracts price candidate fields from fetched snippets/raw content. Amount is discarded if the cited price excerpt is not found in source content. Technical acceptance remains manual.
6. **Price-library accept:** selected quote has positive amount, valid currency/unit/source URL and explicit `Teknik Uyum=Uygun`; normalized source/target unit price stored with URL, VAT/shipping and review metadata.
7. **DXF layer mapping:** semantic auto mapping + user-edited layer/item/measurement/WBS mapping. Metre conversion requires a recognised `$INSUNITS`. HATCH area preferred; closed polyline area fallback. Linear measurement prefers polyline over component LINEs to reduce duplication. Count mode is explicitly labeled as symbol/entity count.

## Bilinen teknik sınırlar (açık)

- **PDF AI ≠ deterministic CAD takeoff.** It interprets what is visible and can miss dimensions, notations, overlaps and revisions. 24 pages maximum per analysis; run separate packages for a large set. A valid evidence string and arithmetic-consistent formula do not certify that AI interpreted the drawing correctly.
- **Donatı ve beton.** Rebar tonnage requires bar bending schedule or explicit diameters/lengths/counts and standards. Concrete/earthwork volumes require sections/heights/thicknesses. The app deliberately does not infer these from a 2D area alone.
- **IFC.** Qto fields only. If a model lacks Qto, this code does not calculate arbitrary solids/Boolean geometry.
- **DWG.** Requires ODA converter in a separate isolated worker; disabled by default. DXF manual layer mapping is supported.
- **Web price.** Tavily search finds indexed web sources, not private dealer portals or guaranteed offers. Product match, regional delivery, minimum order, quote validity, VAT and shipping need human check. No current internet pricing is generated without a real Tavily key and reachable API.
- **Official price.** YFK monthly PDF sync remains separately available; official public rayic and a private-brand/market offer are different kinds of data.
- **Currency/VAT.** This version still does not certify a single normalized VAT-inclusive/exclusive project finance basis across every module.

## Acceptance gate

Before production: install dependencies and run preflight; configure PostgreSQL, ClamAV, `OPENAI_API_KEY`, `OPENAI_MODEL`, `TAVILY_API_KEY`, persistent file storage/backups, ODA worker if DWG needed; then process a real project package containing structural, architectural and MEP plans, specification/BOQ and live supplier price samples. Compare quantities against independent engineer takeoff, compare prices against dated supplier quotations, test user permissions/concurrency, and only then baseline the approved budget.

## Honest status

Automated tests passed as specified in `TEST_REPORT.md`. Real external service integration and actual production deployment were not verified in this runtime; do not treat the package as a certified zero-error live service until the acceptance gate passes.


## V15.4.1 ek bulgu — Excel fiyat provenansı
V15.4.0 Excel import akışında boş fiyat yılı güncel yıla, boş kaynak ise dosya adına tamamlanıyordu. Bu davranış teyitsiz eski fiyatı güncelmiş gibi gösterebilirdi. V15.4.1 bunu düzeltti: açık kaynak ve yıl yoksa fiyat boş kalır, eski tutar yalnızca açıklamada referans olarak tutulur. 6 yeni regresyon testi eklendi.
