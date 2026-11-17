# V15.5.0 Gelişmiş Nihai Sürüm Teknik Denetimi

**Tarih:** 16 Kasım 2026  
**Denetim tabanı:** Emir Asel İnşaat V15.4.1  
**Sürüm:** V15.5.0  
**Değerlendirme:** Test kapsamındaki kod doğrulandı; üretim/live kabul şartları henüz karşılanmadı.

## 1. Ne geliştirildi

### A. Çoklu sağlayıcılı AI katmanı

`ai_gateway.py`, uygulamanın önceki tek sağlayıcıya bağlı AI çağrılarını ortak arayüze taşır:
- OpenAI: Responses API.
- Anthropic: Messages API.
- Google Gemini: Interactions API (`/v1/interactions`); görsel girdiler `image` content olarak dönüştürülür, anahtar `x-goog-api-key` header'ında taşınır.
- OpenAI uyumlu modeller: Chat Completions; base URL, key ve model birlikte zorunlu.

Sağlayıcı ve model kimlikleri yönetici sunucu Secrets alanından gelir. Uygulama kullanıcı adına sağlayıcı hesabı açmaz, API kredisi oluşturmaz ve API erişimi bulunan varsayılan bir model varsaymaz. ChatGPT web planı ile API erişimi birbirinden ayrıdır. `AI_FALLBACK_ENABLED=0` varsayılanı veri paylaşımını sınırlar. Çoklu AI uzman denetimi açık kullanıcı izni olmadan tetiklenmez.

### B. AI ve deterministik kontrol ayrımı

- Kural motoru düzenli alan/tarih/tutar/WBS/hakediş/iş akışı hatalarını tekrarlanabilir şekilde bulur.
- Çoklu AI uzman denetimi; teknik ofis-metraj, maliyet, CPM/EVM, satın alma/sözleşme, saha/İSG ve finans/nakit/hakediş rollerinden gelen bulguları sağlayıcı bazında ayrı raporlar.
- AI girdisinde bütçe, gerçekleşen, program, risk, nakit, satın alma, işçilik, stok, sözleşmeler, değişiklikler, taşeron/işveren hakedişleri, saha günlükleri, onaylar, metraj adayları, piyasa teklifleri ve baseline/snapshot verileri bulunabilir.
- AI çıktısı kaynak veriyi değiştirmez; onay, satın alma, ödeme, bütçe veya hakediş işlemi yapmaz.

### C. Operasyon Kontrol Merkezi

`operations_monitor.py` salt-okunur ek kontrolleri risk, sözleşme, değişiklik emri, taşeron hakedişi, işveren tahsilatı, onay bekleme, nakit vadesi/mükerrerlik, günlük saha raporu ve işçilik tutarları için üretir. Denetim bulgusu hata/eksik veri işaretidir; kaynak satıra otomatik müdahale değildir.

## 2. Test sonuçları

- Python derleme: geçti.
- pytest: 90 geçti, 0 başarısız.
- UI shell: 34 sayfa geçti; Streamlit mock/stub kullanıldı, gerçek tarayıcı testi değildir.
- Sentetik E2E: 38 geçti, 0 başarısız; gerçek IFC ve canlı AI API için 2 harici entegrasyon testi atlandı.
- Gerçek AI API sağlayıcıları için anahtar olmadığı için ağ üzerinden servis denemesi yapılmadı.

## 3. Ünzüle Hanbaba projesiyle ilgili korunan mühendislik/bütçe uyarıları

V15.4.1 proje denetiminde kayıtlı kaynak veriler yeniden kabul edilmeden otomatik olarak onaylı bütçe veya kesin metraj sayılmaz:
- Hiyerarşik bütçe sayfasında 160 WBS satırı Taslak ve fiyat/tutar alanları boştu.
- Çapraz metraj dosyasından 13 satır alınmış; 12 fiyatın kaynağı/yılı eksik olduğundan bütçe fiyatı olarak reddedilmişti.
- Statik alan 199,7 m²; HVAC dokümanında 251,20 m² görünüyordu; alan tanımı gereklidir.
- Donatı grup değerlerinde toplama/parent-child ilişkisi ve çatı geçici çapraz/kalıcı çelik kapsamı mühendis tarafından doğrulanmalıdır.
- Dört DWG AC1032 başlıklıydı; ODA converter olmadığı için gerçek DWG geometri/metraj kontrolü yapılmadı.

## 4. Üretim için açık kalanlar

- Gerçek Streamlit çalışma ortamı/Cloud URL kabulü ve browser testi.
- PostgreSQL, migration, concurrency, backup/restore tatbikatı, kalıcı dosya saklama.
- `EA_ENV=production`, güçlü bootstrap token, ClamAV ve fail-closed upload davranışı.
- ODA converter'ın izole worker içinde kurulup gerçek DWG ile test edilmesi.
- En az bir (çoklu denetim için en az iki) gerçek AI API key/model ile bağlantı, maliyet, rate limit, timeout, provider kullanım şartları.
- Tavily canlı fiyat kaynağı, satıcı/ürün/teknik varyant/KDV/nakliye/bölge/geçerlilik teyidi.
- Gerçek kullanıcı, yetki, eş zamanlı yazma ve onay kilidi kabul testi.

Bu geliştirme, tüm açık çevresel/harici bağımlılıkları kendi başına sağlayamaz. Gerçek API anahtarları, veritabanı, antivirüs, CAD worker ve canlı uygulama adresi olmadan bunlar dürüstçe doğrulanamaz.

## 5. Kabul kararı

**V15.5.0, otomatik ve sentetik kabul testlerini geçen gelişmiş sürüm adayıdır. Henüz canlı üretim kabulü verilmemiştir.** Bu rapor sıfır hata veya mühendislikte kesinlik garantisi değildir. Proje miktar/fiyat/teknik onayı ilgili yetkili profesyonellerde kalır.
