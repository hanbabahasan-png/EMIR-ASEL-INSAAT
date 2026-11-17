# EMİR ASEL İNŞAAT — Entegre Proje Kontrol ERP V15.5.1

**Teslim türü:** Kaynak kodu, kurulum ayarları, testleri, sürüm yedekleri ve sentetik proje kabul dosyalarını içeren geliştirilmiş sürüm adayı. Paket gerçek proje dosyası, müşteri verisi, API anahtarı veya veritabanı parolası içermez.

## Ana hedef ve iş akışı

**Proje dosyası → imalat/teknik tarif tanıma → kanıtlı metraj adayı → WBS/poz ve birim eşleştirme → resmi rayiç/yerel kütüphane/web fiyat araştırması → uzman kontrolü → bütçe taslağı → bütçe onayı/baseline → satın alma/fatura/gerçekleşen → CPM/EVM/nakit akışı/hakediş/risk ve portföy yönetimi.**

AI/CAD/BIM sonuçları taslak ve karar desteğidir. Teknik ya da finansal onay yerine geçmez. Ölçü, kaynak veya teknik uyum belirsizse sistem miktar/fiyat uydurmaz; satırı kontrol bekliyor veya veri eksik olarak bırakır.

## 1. V15.4.1 temel işlevler (V15.5.1 içinde korunur)

### PDF proje metrajı
- Yeni **AI Metraj Çıkarımı ve Piyasa Fiyat Araştırması** ekranı PDF proje paftalarını sayfa ve görsel olarak inceler. Her analizde 1–24 sayfa, 1–4 sayfalık API grupları kullanılır.
- AI'dan disiplin, pafta no, mahal/kat, imalat, teknik tarif, görülen miktar/birim, yöntem, formül, sayfa kanıtı ve güven skoru istenir.
- Metraj sadece açık tablo/Qto veya paftada açıkça bulunan ölçülerden denetlenebilir bir formülle çıkarılabilir. Piksel ölçeğinden, varsayılan kalınlıktan, tipik uygulamadan veya eksik ölçüden miktar tahmini kabul edilmez.
- Dimension calculation formülleri uygulama tarafından güvenli aritmetik gramerle tekrar hesaplanır. Formül ve AI miktarı uyuşmuyorsa miktar boş bırakılır.
- Sayfa, kaynak dosya SHA256, kanıt, yöntem, formül ve güven skoru kayıtla birlikte tutulur.
- Kullanıcı metraj adayını düzenleyebilir, WBS girebilir, yerel fiyat önerisini açıkça kabul edebilir ve yalnızca kontrol bekleyen bütçe taslağına aktarabilir. Sistem bütçe onayı vermez.

### DXF/DWG/CAD metraj
- Bilinen isimli katmanlarda otomatik alan/uzunluk adayları üretilir.
- Yeni **manuel katman eşleştirme** ile standart dışı DXF katmanları imalata ve ölçüm türüne (alan, uzunluk, adet) kullanıcı tarafından bağlanabilir.
- Bilinen birim dönüşümü yoksa ölçüm yapılmaz. Alan için HATCH önceliği, kapalı polyline yedeği; uzunlukta polyline önceliği uygulanır. Polyline varken bileşen çizgileri yeniden toplanarak çift sayım yapılmaz.
- Adet sayımı yalnızca çizim sembol/eleman adedidir; gerçek malzeme adedi olduğu uzman tarafından doğrulanmalıdır.
- 2D çizimden doğrudan beton hacmi, hafriyat hacmi veya donatı tonajı uydurulmaz. Donatı tonajı için donatı listesi/çap-boy-adet-büküm-bindirme verileri gerekir.
- DWG için ODA File Converter ayrıca kurulmalı; üretimde ayrı ve kısıtlı worker ortamında açılmalıdır. Varsayılan kapalıdır.

### IFC/BIM
- IFC'deki açık Qto alanları, tanınan eleman türleri ve proje ölçü birimi dönüşümü üzerinden aday yapılır.
- Qto bulunmayan ürünler için genel amaçlı geometrik quantity takeoff yapılmaz. Bu nedenle IFC modelinin Qto veri kapsamı ve koordinat/üniteleri doğrulanmalıdır.

### İnternet fiyat araştırması
- Tavily Search API ile Türkiye/yerel bölgeye göre web kaynakları araştırılır. Resmî API uç noktası `https://api.tavily.com/search`; arama için `TAVILY_API_KEY` gerekir. Tavily arama sonucundaki URL ve içerik korunur.
- Yapılandırılmış AI gateway (seçilen sağlayıcı/model) arama kaynaklarının ham metninden açıkça belirtilen fiyat, kaynak birimi, para birimi, KDV/nakliye durumu ve ürün açıklaması adayını çıkarır.
- Fiyat ifadesi kaynak metinde doğrulanmıyorsa fiyat hücresi boş tutulur. Kaynak bulunsa bile sayfada fiyat görünmüyorsa kayıt yalnızca kaynak inceleme adayıdır.
- İnternet araması fiyat teklifi/tedarik taahhüdü sayılmaz. Ürün kodu, teknik sınıf, ölçü, fiyat birimi, KDV, nakliye, bölge, minimum sipariş ve geçerlilik tedarikçiyle doğrulanmalıdır.
- Teknik uygunluğu kullanıcı açıkça `Uygun` olarak onaylamadan internet fiyatı şirket fiyat kütüphanesine girmez. Kabul edilen teklif kaynak URL'si, kontrol tarihi ve KDV/nakliye bilgisiyle saklanır.
- Resmî YFK aylık rayiç ve birim fiyatları ayrı PDF senkronizasyon modülüyle işlenir; YFK rayici özel marka/ürün tedarikçi fiyatı yerine geçmez.

### Fiyat eşleştirme ve bütçeye aktarma
- Birim, para birimi, teknik tarif ve imalat kapsamı eşleşmesi zorunludur.
- Kaynak/yıl bilgisi olmayan fiyat önerisi bütçeye kabul edilmez. Eski fiyat işaretli öneri otomatik kabul edilemez.
- AI metraj satırı yalnızca kullanıcı `Yerel Fiyatı Kabul Et` kutusunu işaretleyip `Bütçeye Aktar` seçerse fiyatla bütçe taslağına geçer. Aksi halde fiyat boş kalır.
- Onaylı bütçe ve baseline yine ayrı yetki/onay akışından geçer.

### Şirket portföyü ve maliyet kontrolü
- Proje erişim yetkisi ve para birimi korunur; kur politikası olmadan TRY/EUR/USD/GBP tek tutara toplanmaz.
- Satın alma taahhütleri, fatura ve onaylı gerçekleşen zincirinde mutabakat kontrolü yapılır. Belirsiz fatura eşleşmesi taahhütten düşülmez.

## 2. Canlı hizmet / yapılandırma

`.env` veya Streamlit Secrets içine gerekli ayarları girin. AI belge/metraj analizi için aşağıdaki sağlayıcılardan en az birine ait **API anahtarı + hesabınızda etkin model kimliği** gerekir. Çoklu AI çapraz denetimi için en az iki farklı sağlayıcı yapılandırın. İnternet fiyat araştırması için ayrıca **TAVILY_API_KEY** gerekir. API anahtarı veya model yoksa AI özellikleri kapalı kalır; fiyat/analiz varmış gibi gösterilmez.

```dotenv
EA_ENV=development
EA_BOOTSTRAP_TOKEN=REPLACE_WITH_RANDOM_SECRET_MIN_32_CHARS
EA_BASE_CURRENCY=TRY
AI_PROVIDER=auto
AI_FALLBACK_ENABLED=0
AI_FALLBACK_ORDER=openai,anthropic,gemini,openai_compatible
AI_TIMEOUT_SECONDS=90
AI_MAX_OUTPUT_TOKENS=4096
AI_REVIEW_MAX_PROVIDERS=3
AI_STORE_RESPONSES=0
OPENAI_API_KEY=
OPENAI_MODEL=gpt-6-astra  # hesabınızda erişim varsa; değilse erişilebilir model kimliğini girin
ANTHROPIC_API_KEY=
ANTHROPIC_MODEL=
GEMINI_API_KEY=
GEMINI_MODEL=
GEMINI_API_VERSION=v1
AI_COMPAT_BASE_URL=
AI_COMPAT_API_KEY=
AI_COMPAT_MODEL=
TAVILY_API_KEY=
DATABASE_URL=
ENABLE_ODA_CONVERSION=0
ODA_FILE_CONVERTER=
CLAMSCAN_PATH=
```

### Yapay zekâ sağlayıcıları ve çalışma ilkeleri
- **OpenAI API (GPT-6):** OpenAI Responses API kullanılır. OpenAI geliştirici hızlı başlangıç sayfasındaki güncel GPT-6 örnek model kimliği `gpt-6-astra` olarak verilmiştir; `OPENAI_MODEL` alanını yalnız hesabınızda bu modele erişim varsa kullanın, aksi durumda erişim izni bulunan tam model kimliğini girin. ChatGPT web aboneliği, API kullanım hakkı veya aynı model kimliğine otomatik erişim vermez; API anahtarı/billing ayrıca yönetilir.
- **Anthropic Claude:** Messages API üzerinden metin ve desteklenen görsel girdileri kullanılır.
- **Google Gemini:** Güncel Interactions API (`/v1/interactions`) üzerinden metin ve base64 görsel içeriği aktarılır; sonuç saklama varsayılan olarak kapalıdır.
- **OpenAI uyumlu özel uç nokta:** `AI_COMPAT_BASE_URL`, `AI_COMPAT_API_KEY` ve `AI_COMPAT_MODEL` birlikte tanımlanmalıdır.
- `AI_PROVIDER=auto` sıralı sağlayıcı seçer. `AI_FALLBACK_ENABLED=0` varsayılanı, başarısızlık halinde başka sağlayıcıya otomatik veri gönderilmesini önler. Yedek sağlayıcıyı açarsanız gönderilen metin/görsel içerik, başarısızlık halinde listedeki diğer sağlayıcıya da iletilebilir.
- Çoklu AI uzman denetimi ayrı bir ekrandır. Kullanıcının açık onayından sonra yapılandırılmış sağlayıcıların her biri aynı sınırlı proje özetini bağımsız inceler. Sonuçlar ayrı gösterilir; otomatik birleştirilmez, proje kayıtlarına yazılmaz veya onaylanmaz.
- API anahtarlarını uygulama koduna, ZIP'e veya Git deposuna yazmayın. Anahtarları yalnız sunucu Secrets/environment alanında saklayın. `AI_STORE_RESPONSES=0` isteğin sağlayıcı tarafında saklanmasını kapatan parametreyi kullanır; her sağlayıcının veri saklama/işleme koşulları ayrıca incelenmelidir.

Üretim modunda `EA_ENV=production` için PostgreSQL, 32+ karakter rastgele bootstrap token ve çalışan ClamAV taraması zorunludur. Kalıcı dosya depolama, yedekleme ve SSL/TLS reverse proxy de sunucu altyapısında kurulmalıdır. API anahtarlarını kaynak koda yazmayın.

## 3. Windows kurulum

Python 3.11+ önerilir.

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
# Test araçları da kurulacaksa:
pip install -r requirements-dev.txt
Copy-Item .env.example .env
python scripts/preflight.py
python -m streamlit run app.py
```

## 4. Testler

```bash
python -m py_compile app.py tests/*.py scripts/preflight.py
pytest -q
python tests/test_ui_shell_smoke.py
python tests/e2e_project_acceptance.py
python scripts/preflight.py
```

V15.5.1 mevcut kaynak + güvenli kurulum yaması doğrulaması: **104 birim/runtime testi geçti**; **34/34 sayfa** taklit Streamlit kabuğunda hatasız açıldı; sentetik proje kabulinde **38 geçti, 0 başarısız, 2 harici entegrasyon testi atlandı**. Bunlar gerçek tarayıcı/servis kabul testi yerine geçmez.

## 5. Paket yardımcıları

- `ai_gateway.py`: OpenAI / Claude / Gemini / uyumlu servis sağlayıcı katmanı.
- `operations_monitor.py`: kural tabanlı proje operasyon kontrolleri.
- `security_config.py`: bootstrap anahtarı placeholder ve HTML escape güvenlik yardımcıları.
- `requirements-dev.txt`: pytest ve coverage gibi geliştirme/test bağımlılıkları.
- `FINAL_AUDIT_REPORT_TR.md`: son teknik denetim, test sınırları ve üretim öncesi açık işler.

## 6. Bilinen sınırlar / canlı kabul

- Bu ortamda OpenAI, Anthropic, Gemini ve Tavily için gerçek API anahtarları sağlanmadı; sağlayıcı uçtan uca canlı sorgu yapılmadı. AI ağ geçidi ve sağlayıcı istek/yanıt dönüşümü mock testlerle doğrulandı.
- Gerçek Streamlit tarayıcı açılışı ve hosted/cloud deploy testi bu ortamda tamamlanmadı.
- IFC test dosyası için `ifcopenshell` kurulu değildi; yalnızca opsiyonel dependency davranışı ve Qto dönüştürme mantığı kontrol edildi.
- DWG gerçek ODA dönüşümü, ClamAV imza veritabanı, PostgreSQL sunucusu ve çok kullanıcılı yük testi canlı altyapıda kabul edilmelidir.
- PDF AI metrajı her imalatı, her disiplini veya her paftayı hatasız çıkarma garantisi vermez. Özellikle demir donatı tonajı, hacim, kazı, boşluk düşümü, bindirme ve fire gibi mühendislik kuralları donatı çizelgesi/hesap/pafta verisi ve uzman kontrolü gerektirir.
- PDF metraj şu an 1–24 sayfa/çalışma ile sınırlıdır; büyük projeler disiplin ve pafta paketlerine bölünerek işlenir.
- Piyasa fiyatı arama internette indekslenmiş kaynakları bulur. Kapalı bayi portalları, üyelik isteyen fiyatlar ve telefonda teklif gerektiren ürünler otomatik okunamaz.
- KDV dahil/hariç/net maliyet standardı bütün ERP modülleri boyunca tam normalize edilmiş kabul edilmemelidir; teklif ve satın alma öncesi aynı vergi/nakliye bazında karşılaştırma yapılmalıdır.
- Bu yazılım proje kontrol karar desteğidir; mühendislik hesap raporu, ihale keşfi, resmi keşif veya sertifikalı BIM takeoff hizmeti yerine geçmez.

Canlı kabule geçmeden önce `docs/ACCEPTANCE_CHECKLIST_TR.md` içindeki gerçek proje senaryolarını uygulayın.


## V15.4.1 fiyat kaynağı güvenlik düzeltmesi
Excel içe aktarımında fiyat yılı veya açık fiyat kaynağı eksikse tutar bütçe birim fiyatına alınmaz; referans açıklamada korunur. İki yıl ve daha eski fiyatlar yeni teklif doğrulamasına bırakılır. Ayrıntı: `RELEASE_NOTES_V15_4_1.md`.


## 6. V15.5.1 yeniden denetim notu (16.11.2026)

Güncel çok dosyalı paket üzerinde testler yeniden çalıştırıldı: 100/100 pytest (14 subtest dahil), 34/34 UI kabuk testi; sentetik E2E’de 38 geçti, 0 başarısız oldu ve IFC dosya entegrasyonu ile canlı AI API çağrısı için 2 harici entegrasyon testi atlandı. Gerçek Streamlit Cloud/üretim tarayıcı testi yapılmadı. Canlı kabul engelleri ve Ünzüle Hanbaba projesine ilişkin metraj/bütçe bulguları `ADVANCED_RELEASE_AUDIT_V15_5_0_TR.md` ve `LIVE_ACCEPTANCE_STATUS_TR.md` içindedir.


## V15.5.1 — Çoklu AI, bütünleşik operasyon kontrolü ve veri bütünlüğü düzeltmeleri

- OpenAI Responses, Anthropic Messages, Google Gemini Interactions ve OpenAI uyumlu API için sağlayıcı-bağımsız ağ geçidi eklendi.
- Otomatik sağlayıcı yedekleme gizlilik için varsayılan kapalıdır; çoklu sağlayıcı incelemesi ayrıca kullanıcı onayı ister.
- Proje asistanındaki AI uzman denetimi bağlamı; bütçe, gerçekleşen, iş programı, risk, nakit, satın alma, işçilik, stok, sözleşme/değişiklik, iki yönlü hakediş, saha günlükleri, onaylar, metraj adayları, piyasa teklifleri ve baseline verilerini kapsayacak şekilde genişletildi.
- Operasyon kontrol merkezi risk, sözleşme, değişiklik emri, taşeron/işveren hakedişi, nakit, onay, saha günlükleri ve işçilik tutarlılığı için deterministik, salt okunur bulgular üretir.
- AI çıktıları öneri olarak kalır; kayıt yazma, bütçe/teknik onay, ödeme, satın alma ve hakediş kararı vermez.


## 7. V15.5.1 veri bütünlüğü ve metraj aktarım düzeltmeleri

- Onaylanmış plan, gerçekleşen, satın alma, sözleşme, değişiklik, taşeron hakediş ve işveren hakediş kaydında iş alanları değişirse kayıt merkezi veritabanı katmanında yeniden `Kontrol Bekliyor` durumuna alınır; onay izleri temizlenir ve değişiklik notu eklenir. Onaylı kayıt doğrudan silinemez.
- `takeoff_candidates` için yayın/kontrol tarihleri ve `Kaynak Sayfa` alanı, `market_quotes` için kaynak SHA256, kaynak dosya ve arama sorgusu eklendi.
- Excel metraj otomatik kolon eşleştirmesi `Kalem` başlığını tanır. Genel `Kaynak Durumu` sütunu fiyat kaynağı gibi yorumlanmaz.
- Ünzüle Hanbaba gerçek metraj çalışma sayfasında 13 satır aday olarak okundu; 8 satır miktar/birim yönünden taslak aktarılabilir bulundu. Açık fiyat yılı/kaynağı bulunmayan 13 satırın hiçbir fiyatı bütçeye aktarılmadı.
- Stok hareketli ortalama maliyeti, ağırlıklı ilerleme ve kalite kontrol motoruna doğrudan testler eklendi.
