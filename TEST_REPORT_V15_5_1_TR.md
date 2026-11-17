# V15.5.1 — Test ve Teknik Kabul Raporu

**Durum: OTOMATİK TESTLER GEÇTİ / CANLI ÜRETİM KABULÜ TAMAMLANMADI**

## 1. Bu oturumda tekrar çalıştırılan testler

| Kontrol | Sonuç | Sınır / açıklama |
|---|---:|---|
| Python derleme (`py_compile`) | GEÇTİ | `app.py`, AI gateway, operasyon modülü ve test modülleri |
| pytest | **105 geçti, 0 başarısız** | Ayrıca 14 subtest başarılı |
| Streamlit UI kabuk testi | **34/34 geçti** | Taklit Streamlit nesnesi; gerçek tarayıcı testi değil |
| Sentetik proje uçtan uca | **38 geçti, 0 başarısız, 2 atlandı** | İzole SQLite ve sentetik villa projesi |
| ZIP bütünlük testi | GEÇTİ | Son V15.5.1 teslim arşivi oluşturuldu ve `unzip -t` ile doğrulandı |
| Python satır kapsamı | **%47 toplam** | `app.py` %44, `ai_gateway.py` %63, `operations_monitor.py` %88 |

İki atlanan E2E testi: gerçek IFC dosyası testi (`ifcopenshell` yok) ve canlı AI API testi (gerçek sağlayıcı anahtarı yok).

## 2. Gerçek Ünzüle Hanbaba proje kaynakları

Kaynak arşivindeki 15 uygulama dosyası (6 PDF, 5 XLSX, 4 DWG) uygulamanın uzantı/imza doğrulama hattından geçti. Beş Excel çalışma kitabı `openpyxl` ile açıldı; altı PDF PyMuPDF ile açılarak sayfa sayısı ve ilk sayfa metin okunabilirliği kontrol edildi. Dört DWG dosyası `AC1032` başlıklı AutoCAD 2018 dosyası olarak tanındı.

**Gerçek DWG geometrisi/metrajı çalıştırılmadı:** ODA File Converter bulunmuyor ve DWG dönüştürme üretim güvenliği için varsayılan kapalı. Bu kontrol yalnız dosya imzasını ve güvenli yükleme doğrulamasını kapsar.

`Villa_Insaat_Proje_Capraz_Metraj_v2.xlsx` dosyasının `Metraj` sayfası gerçek import motoruyla kontrol edildi: 13 aday satır üretildi, 8 satır miktar ve birim yönünden taslak aktarım için uygundu. Kaynak/yıl kanıtı bulunmadığı için **0 fiyat** bütçeye kabul edildi. `Kalem` başlığı otomatik eşleştirme listesine eklendi.

Paket DXF geometrisi 3 entity ve 3 layer ile okundu; geometri parser hata listesi boş döndü. Bu, gerçek DWG değil, test amaçlı DXF paftasıdır.

## 3. V15.5.1'de düzeltilen regresyonlar

- Önceden onaylanmış gerçekleşen/satın alma/sözleşme/değişiklik/hakediş ve bütçe satırında iş alanı değişirse onay artık sessizce korunmaz; kayıt yeniden kontrol/onaya döner. Onaylı satırın silinmesi merkezi kayıt katmanında engellenir.
- `takeoff_candidates` yayın/kontrol tarih alanları ve tarih normalizasyonu eklendi; `Kaynak Sayfa` tutulur.
- `market_quotes` kaynak SHA256, kaynak dosya ve arama sorgusu ile izlenebilir hale getirildi.
- Excel importunda `Kalem` başlığının otomatik eşleştirilmesi eklendi; `Kaynak Durumu` gibi durum alanları birim fiyat kaynağı olarak varsayılmaz.
- Stok ağırlıklı ortalama maliyet, negatif stok uyarısı, ağırlıklı ilerleme ve kalite kontrol kontrolleri için regresyon testleri eklendi.

## 4. Canlı ortamda test edilemeyenler / engeller

| Alan | Durum | Neden |
|---|---|---|
| Gerçek Streamlit başlangıcı / browser / staging deploy | BLOKE | Bu çalışma ortamında `streamlit` kurulu değil; pip kurulum denemesi DNS/ağ erişimi nedeniyle başarısız oldu. Kullanıcının güncel deploy URL'si de bu test turunda verilmedi. |
| OpenAI / Claude / Gemini canlı API | ATLANDI | API anahtarları ve model erişimleri yok; adapter testleri mock yanıtlarla sınırlı. |
| YFK canlı senkronizasyon | ATLANDI | Uygulama içinden resmî siteye gerçek ağ çağrısı yapılmadı. |
| Tavily piyasa fiyat araştırması | ATLANDI | `TAVILY_API_KEY` tanımlı değil. |
| DWG→DXF geometri dönüştürme | BLOKE | ODA File Converter kurulu/yapılandırılmış değil. |
| Gerçek IFC Qto okuma | ATLANDI | `ifcopenshell` kurulu değil ve arşivde kabul için kullanılabilir IFC dosyası yok. |
| ClamAV / zararlı yazılım taraması | ATLANDI | `CLAMSCAN_PATH` ve imza veritabanı yok. Dosyalar yalnız format/imza kontrolünden geçti; bu, antivirüs taraması değildir. |
| PostgreSQL advisory-lock / iki kullanıcı eşzamanlı yazma | ATLANDI | PostgreSQL servisi ve `psycopg` sürücüsü mevcut değil. SQLite çakışma/optimistic concurrency testleri canlı PostgreSQL testi değildir. |
| Gerçek yedek geri yükleme tatbikatı | ATLANDI | Uygulama JSON yedeği veri setlerini kapsar; dosya varlıkları, kullanıcı tabloları ve veritabanının tamamı için `pg_dump` + dosya yedeği restore tatbikatı yapılmadı. |

## 5. Hâlâ üretim öncesi çözülmesi gereken ürün/güvenlik işleri

1. **KDV esası:** Gerçekleşen/fatura maliyetlerinde KDV dahil/hariç tutarları ayrı saklayan ve bütçeyle aynı maliyet esasına çeviren veri modeli tamamlanmadı. Ekran mevcut sürümde bu uyumsuzluk için uyarı verir; otomatik KDV ayrıştırması yapmaz.
2. **DWG otomatik metrajı:** Şu an gerçek DWG geometri okuması yoktur. Uygulanabilir yol ODA dönüşüm worker'ı veya DXF yarı otomatik akışıdır; katman–imalat eşleştirmesi, birim ve kalınlık kullanıcı/mühendis onayına bağlı kalır.
3. **Gerçek servis kabulü:** Streamlit staging, canlı AI, Tavily, YFK, PostgreSQL eşzamanlılık, ClamAV ve IFC kabul testleri yapılmalıdır.
4. **KVKK/veri yaşam döngüsü:** Proje kapatıldığında/silindiğinde dosya ve verilerin saklama-silme süresi ile hukuki imha akışı için yazılı ve uygulanmış politika tamamlanmadı.
5. **MFA/SSO ve erişim sertleştirme:** Rol bazlı erişim mevcuttur; kurumsal MFA/SSO ve kapsamlı servis/API rate limiting canlı ortamda ayrıca tamamlanmalıdır.
6. **Modülerleştirme:** `app.py` 5.000 satırın üzerindedir ancak dosya kesik değildir; derlenir ve son UI handler'a kadar tamamlanır. Bakım kolaylığı için sonraki refaktörde UI, veri erişimi, CAD, güvenlik ve YFK modüllerine ayrılması önerilir.

## 6. Sonuç

Bu rapordaki testler kodun mevcut otomatik test kapsamındaki davranışını doğrular; **sıfır hata, her kod yolunun test edildiği, gerçek çizim metrajının doğru olduğu veya üretim sisteminin hazır olduğu anlamına gelmez**. Ölçülen toplam satır kapsamı %47'dir. Canlı entegrasyon ve güvenlik koşulları kapanmadan V15.5.1 üretim kabulü verilmemelidir.
