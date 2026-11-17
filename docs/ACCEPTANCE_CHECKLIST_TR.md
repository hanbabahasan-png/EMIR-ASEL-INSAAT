# V15.4.1 Canlı Kabul Kontrol Listesi

## A. Altyapı
- [ ] Python 3.11+ ve requirements kurulumu
- [ ] `python scripts/preflight.py` hatasız
- [ ] Streamlit browser açılışı ve mobil/desktop temel kontrol
- [ ] PostgreSQL kalıcılık, backup/restore ve iki kullanıcılı eşzamanlı kayıt kontrolü
- [ ] ClamAV imza DB + EICAR güvenli test senaryosu; tarama başarısızken yüklemenin bloke olması
- [ ] Secrets: `OPENAI_API_KEY`, `OPENAI_MODEL`, `TAVILY_API_KEY`; secrets values not logged

## B. Gerçek proje okuma/metraj
- [ ] Mimari, statik, mekanik ve elektrik paftaları ayrı PDF gruplarında yükle
- [ ] PDF sayfa no, pafta no, mahal, imalat, teknik tarif, kanıt ve formül doğrula
- [ ] Direct table quantity and dimension-calculated quantities independently re-measure
- [ ] Deliberately mismatched arithmetic formula remains quantity blank
- [ ] Low-confidence/evidence missing rows cannot be transferred
- [ ] Duplicate plan/revision pages do not double count
- [ ] Rebar quantity is compared with BBS, not accepted from 2D area

## C. CAD/BIM
- [ ] DXF units ($INSUNITS) and origin/scale verified
- [ ] Semantic auto-layer candidates checked against CAD layers
- [ ] Manual layer mapping tested for area/length/count
- [ ] Closed polyline/HATCH overlap and hidden/external references checked
- [ ] DWG ODA worker isolation, timeout and output validation
- [ ] IFC units, Qto set coverage, global IDs and storey assignments checked

## D. Price research
- [ ] Tavily search returns current supplier/producer/public source links
- [ ] Source raw content visibly contains the claimed amount and unit
- [ ] Exact product class/thickness/diameter/standard/brand/model checked
- [ ] currency, quote date, regional delivery, MOQ, VAT/shipping and validity checked
- [ ] Different unit conversion verified independently
- [ ] Technical suitability is explicitly accepted by a responsible user
- [ ] YFK official rayic is distinguished from real supplier offer
- [ ] Old prices are blocked/researched; no price created when source is missing

## E. Budget/actual and project controls
- [ ] Metraj quantities tie to unique WBS/position/measurement unit
- [ ] All price and currency assumptions reviewed before budget approval
- [ ] Budget baseline approval workflow completed
- [ ] Purchase order/invoice/actual chain and duplicate invoice behavior tested
- [ ] VAT basis and shipping costs compared on the same basis
- [ ] CPM/EVM/cash flow/hakediş/dashboard results reconciled to source data
- [ ] Audit trail and Excel output checked

## Live limitations to acknowledge
AI PDF reading is an assistant, not an independent engineer certification. IFC currently uses existing Qto quantities only. 2D DXF area/length is not equivalent to full BOQ; concrete volume/rebar tonnage needs additional source dimensions/schedules. Web results are market-research leads, not binding supplier offers. Do not baseline or purchase based solely on automatic output.
