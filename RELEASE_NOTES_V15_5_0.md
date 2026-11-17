# Emir Asel İnşaat — V15.5.0 Sürüm Notları

**Tür:** V15.4.1 üzerine gelişmiş çoklu AI / operasyon kontrol sürümü
**Tarih:** 16.11.2026

## Yeni
- OpenAI Responses API, Anthropic Messages API, Google Gemini Interactions API ve OpenAI uyumlu Chat Completions için ortak AI gateway.
- Sağlayıcı/model yapılandırma tablosu; otomatik fallback varsayılan kapalı.
- İzinle başlatılan bağımsız çoklu AI uzman denetimi: teknik ofis/metraj, bütçe-maliyet, CPM/EVM, satın alma/sözleşme, saha/İSG, finans/nakit/hakediş.
- AI analiz bağlamına gerçekleşen, nakit, satın alma, işçilik, stok, saha günlükleri, sözleşmeler, değişiklikler, iki yönlü hakediş, onaylar, metraj adayları, piyasa teklifleri ve baseline/snapshot verileri dahil edildi.
- Operasyon kontrol merkezi için genişletilmiş salt-okunur deterministik kontroller.
- AI proje bağlamı uzun olduğunda geçerli JSON biçimini koruyarak kanıt metinleri/kayıtları kısaltan güvenli bağlam sıkıştırması.
- Gemini API key URL query yerine `x-goog-api-key` header üzerinden iletilir; API isteğinde response storage varsayılan kapalıdır.

## Güvenlik ve sınırlar
- AI bulguları yalnız öneridir, insan/uzman onayı bekler. Hiçbir AI sonucu otomatik bütçe, gerçekleşen, hakediş, satın alma, ödeme, stok veya sözleşme değişikliği yapmaz.
- Çoklu AI denetimi tüm hazır sağlayıcılara aynı sınırlı proje özetini gönderir; her çalıştırma için kullanıcı onayı gerekir.
- Otomatik fallback kapalıdır; açılırsa başarısız sağlayıcı sonrası aynı içerik başka bir sağlayıcıya gönderilebilir.
- Model isimleri API hesabına göre değiştiği için örnek dosyada boş bırakılmıştır; kullanıcı exact API model ID tanımlar. ChatGPT web planı API kredisi/erişimi sağlamaz.

## Test sonucu
- pytest: 90 passed.
- UI kabuk testi: 34/34 passed; gerçek tarayıcı testi değildir.
- Sentetik proje E2E: 38 passed, 0 failed, 2 skipped (gerçek IFC dosyası ve canlı AI API; bağımlılık/anahtar yok).
- Canlı Streamlit/hosted URL kabul testi yapılmadı. `streamlit` paketi bu çalışma ortamında yok; PostgreSQL, ClamAV, ODA, Tavily ve sağlayıcı API anahtarları da yapılandırılmadı.
