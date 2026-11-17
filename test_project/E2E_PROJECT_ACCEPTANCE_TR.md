# Adana Villa E2E — Proje Kabul Testi

- Tarih: 16.11.2026 15:28
- Test verisi: Sentetik örnek proje; gerçek müşteri projesi değildir.
- Sonuç: **38 geçti, 0 başarısız, 2 entegrasyon testi çalıştırılamadı**

## Test detayları

- [PASS] Test DB ilk kurulum ve yönetici kimlik doğrulama
- [PASS] ISO ve Türkçe tarih biçimlerinin uyarısız normalizasyonu
- [PASS] Proje kartı oluşturma ve proje erişim ataması
- [PASS] Gerçek uygulama başlangıç akışı / yönetici paneli kabuğu — headless Streamlit stub; gerçek tarayıcı değil
- [PASS] Örnek proje Excel dosyasını güvenli doküman deposuna yükleme — 1 belge kaydı
- [PASS] Aynı proje dosyasının mükerrer yüklenmesini engelleme
- [PASS] Excel fatura satırının gerçekleşen adayı olarak okunması
- [PASS] Excel ekranındaki ortak metraj hazırlama motoru ve WBS alanı
- [PASS] Excel metraj aktarımında WBS kodlarının korunması
- [PASS] WBS kodlarının büyük harfe normalize edilmesi
- [PASS] Excel birim fiyatı, yılı ve kaynak bilgisinin korunması
- [PASS] Excel içinde yinelenen WBS kodlarının engellenmesi
- [PASS] Projede mevcut WBS kodlarının mükerrer importunun engellenmesi
- [PASS] Excel metraj sayfasının WBS ve imalat sınıflarıyla okunması
- [PASS] İş programı sayfasının bütçe WBS kodlarıyla okunması
- [PASS] Bütçe tutarlarının hesaplanması — Toplam 2,514,600.00 TRY
- [PASS] Kaba / ince / cephe / mekanik / elektrik / yalıtım sınıflandırması — {'KAB.01': ('Kaba İşler', 'Grobeton'), 'KAB.02': ('Kaba İşler', 'Temel / Radye Betonu'), 'KAB.03': ('Kaba İşler', 'Donatı / İnşaat Demiri'), 'YAL.01': ('Yalıtım İşleri', 'Temel / Perde Su Yalıtımı'), 'INC.01': ('İnce İşler', 'Seramik / Fayans'), 'CEP.01': ('Cephe İşleri', 'Cephe Kaplama'), 'MEK.01': ('Mekanik Tesisat', 'Temiz / Pis Su Tesisatı'), 'ELE.01': ('Elektrik Tesisat', 'Kuvvetli Akım')}
- [PASS] Gerçek fiyat kütüphanesi eşleştirmesi — YÜKSEK EŞLEŞME — ONAY GEREKLİ
- [PASS] Bütçe satırlarının çoklu onay talebine gönderilmesi
- [PASS] Yönetici tarafından bütçe onayı ve denetim izi
- [PASS] Onaylı bütçe + iş programı baseline oluşturma ve hash doğrulama
- [PASS] Demir faturası: 5.000 kg → 5 ton ve toplam tutar korunması — {'WBS': 'KAB.03', 'Birim': 'ton', 'Gerçekleşen Miktar': 5.0, 'Gerçekleşen Birim Fiyat': 34500.0, 'Gerçekleşen Tutar': 172500.0}
- [PASS] Gerçekleşen onaylanmadan EVM AC dışı kalması
- [PASS] Gerçekleşen onayından sonra AC hesabına dahil edilmesi — AC=172500.0
- [PASS] WBS bazında bütçe/gerçekleşen demir karşılaştırması — {'Bütçe Miktarı': 22.0, 'Gerçekleşen Miktar': 5.0, 'Bütçe Tutarı': 748000.0, 'Gerçekleşen Tutar': 172500.0, 'Teknik Kontrol': 'UYGUN'}
- [PASS] Ana imalat grubu yönetim özeti üretilmesi
- [PASS] Otomatik proje kontrol bulguları üretilmesi
- [PASS] Nakit akışı kaydı saklama ve yükleme
- [PASS] İşveren hakediş kesinti/KDV/tahsilat hesabı — {'Net Tahakkuk': 85000.0, 'KDV Tutarı': 17000.0, 'Hakediş Toplamı': 102000.0, 'Kalan Alacak': 52000.0}
- [PASS] DXF dosya yükleme ve geometri okuma
- [PASS] DXF ölçekli alan hesabı (mm → m²) — test alanı yaklaşık 80,00 m²
- [PASS] DXF katmanından seramik alan metraj adayı üretimi — [{'Birim': 'm²', 'Miktar': 30.0}]
- [PASS] DXF aynı dosya/katman tekrar aktarım kontrolü
- [PASS] DWG dönüştürücü yapılandırılmadığında güvenli durması
- [PASS] IFC opsiyonel bağımlılık güvenli hata kontrolü — IFC okumak için opsiyonel ifcopenshell kurulmalı.
- [SKIP] Gerçek IFC dosyası uçtan uca entegrasyon testi — ifcopenshell bu test ortamında kurulu değil; gerçek IFC dosyası doğrulanamadı
- [PASS] AI sağlayıcı anahtar/model yapılandırma kontrolü — Yapay zekâ sağlayıcısı yapılandırılmamış. OPENAI_API_KEY + OPENAI_MODEL veya ANTHROPIC_API_KEY + ANTHROPIC_MODEL ya da GEMINI_API_KEY + GEMINI_MODEL tanımlayın.
- [SKIP] Canlı AI API uçtan uca testi — OpenAI/Claude/Gemini anahtarları bu test ortamında tanımlı değil; canlı AI çağrısı yapılmadı
- [PASS] Dolu test projesiyle tüm 33 sayfa yönlendirme/render duman testi — 33 sayfa
- [PASS] Excel yönetim raporu çıktısı

## Sınırlar

- Testler gerçek uygulama Python iş fonksiyonları, izole SQLite veritabanı ve taklit Streamlit kabuğu üzerinde çalıştırıldı. Gerçek tarayıcı ve Streamlit Cloud testi değildir.
- DXF dosyası geometrik test fikstürüdür; gerçek mimari/statik pafta değildir.
- DWG dönüştürmesi ODA File Converter bulunmadığı için çalıştırılmadı; güvenli varsayılan kapalılık test edildi.
- ifcopenshell bu ortamda kurulu olmadığından gerçek IFC dosyası işlenemedi.
- OpenAI/Anthropic/Gemini API anahtarları bulunmadığından canlı AI API isteği yapılmadı.
- YFK canlı HTTP/PDF fiyat senkronizasyonu internet erişimi ve güncel kaynağa bağlı olduğundan bu kabul testinde çalıştırılmadı.
- KDV dahil/hariç alan ayrımı ürün kapsamı açığı olarak kaldı.