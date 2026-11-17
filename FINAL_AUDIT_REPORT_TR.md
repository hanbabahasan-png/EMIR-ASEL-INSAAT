# EMİR ASEL İNŞAAT V15.5.1 — Son Teknik Denetim ve QA Raporu

**Yayın kararı: ÜRETİM ONAYI VERİLMEDİ — güvenlik düzeltmeleri uygulanmış QA adayı.**

Bu rapor mevcut V15.5.1 kaynak kodu ve ZIP teslim kopyası üzerinde yapılan tekrar denetimini kapsar. Sürüm numarası artırılmadı; gerçek canlı entegrasyon testleri tamamlanmadan yeni sürüm/üretim kabulü ilan edilmedi.

## 1. Kaynak ve paket doğrulaması

- Bağımsız kaynak dosyası ile teslim ZIP'indeki `app.py` SHA256 değerleri denetim başlangıcında aynıydı.
- `app.py` dosyası tam; Python AST/derleme kontrolü başarılı.
- `app.py` içindeki sayfa yönlendirmeleri 34 sayfa olarak doğrulandı.
- ZIP paketindeki önceden birikmiş `__pycache__` ve `.pytest_cache` klasörleri temizlendi; teslim paketi kaynak ve test dosyalarını içerir.
- `security_config.py` eklendi; uygulama ve preflight aynı güvenlik doğrulamasını kullanır.

## 2. Bu denetimde bulunan ve düzeltilen güvenlik açıkları

### 2.1 İlk yönetici kurulum anahtarı placeholder açığı

`.streamlit/secrets.toml.example` içindeki `BURAYA_32_KARAKTERDEN_UZUN_RASTGELE_TOKEN` değeri uzunluk kontrolünü geçebiliyordu. `auth_gate()` yalnız eski `.env.example` placeholder'ını engelliyordu; preflight da aynı eksik kontrole sahipti. Bu, örnek ayar kopyalanıp değiştirilmezse ilk yönetici hesabının öngörülebilir bir token ile kurulmasına yol açabilirdi.

**Düzeltme:** `security_config.is_placeholder_secret()` eklendi. Uygulama ve preflight bu ortak fonksiyonu çağırıyor. Secrets örneğinde kullanılabilir token değeri boş bırakıldı; placeholder örnekler testle reddediliyor.

### 2.2 Yönetici panelinde HTML enjeksiyonu

Seçili proje adı `unsafe_allow_html=True` kullanılan yönetici kartına doğrudan ekleniyordu. Proje adı kullanıcı girdisi olduğundan HTML enjeksiyonu riski vardı.

**Düzeltme:** `escape_html_text()` eklendi ve proje adı HTML'e yerleştirilmeden önce escape ediliyor. Kötü niyetli `<img onerror=...>` örneğiyle regresyon testi eklendi.

### 2.3 E2E test kopyalama bağımlılığı

Yeni ortak güvenlik modülü eklendikten sonra E2E geçici uygulama klasörüne kopyalanmıyordu. Test düzeneği güncellendi; runtime modüllerinin tümü birlikte kopyalanıyor.

## 3. Tekrar çalıştırılan otomatik testler

| Kontrol | Sonuç | Açıklama |
|---|---:|---|
| `py_compile` | GEÇTİ | app, AI gateway, operasyon, güvenlik, preflight ve test dosyaları |
| `pytest -q` | **105 geçti, 0 başarısız** | Ayrıca 14 subtest başarılı |
| UI kabuk testi | **34/34 geçti** | Taklit Streamlit nesnesi; gerçek tarayıcı testi değil |
| Sentetik proje E2E | **38 geçti, 0 başarısız, 2 atlandı** | Gerçek SQLite, test projesi, dosya yükleme, WBS, onay, baseline, EVM, DXF ve Excel raporu |
| ZIP testi | GEÇTİ | Paket temizlendikten sonra yeniden doğrulandı |

Atlanan iki E2E testi gerçek IFC dosyası/`ifcopenshell` ve canlı AI API anahtarları gerektirir.

## 4. Ünzüle Hanbaba gerçek proje dosyaları

Kaynak ZIP'i açılarak aşağıdaki dosyalar biçimsel olarak incelendi:

- 5 XLSX: tamamı `openpyxl` ile açıldı ve sayfa adları okundu.
- 6 PDF: tamamı PyMuPDF ile açıldı; sayfa sayıları ve ilk sayfa metin uzunlukları okundu.
- 4 DWG: tümü `AC1032` (AutoCAD 2018 ailesi) imzasıyla tanındı.

Bu, dosya okunabilirlik/imza kontrolüdür; DWG geometrisi veya mühendislik metrajı doğrulaması değildir. Dört DWG için ODA File Converter yapılandırılmadığından gerçek geometri dönüşümü yapılmadı. Bazı PDF'lerde ilk sayfa metni çok az/boş olduğundan taranmış sayfa/OCR veya görsel AI analizi gerekir.

### Gerçek metraj Excel'i

`Villa_Insaat_Proje_Capraz_Metraj_v2.xlsx` / `Metraj` sayfası uygulamanın ortak metraj hazırlama motoruyla yeniden işlendi:

- 13 kaynak satırı okundu.
- 13 metraj adayı oluşturuldu.
- 8 satırda pozitif miktar ve birim birlikte mevcut.
- 5 satırda miktar eksik.
- Kaynak ve fiyat yılı olmadığı için 13 satırın hiçbir fiyatı bütçeye kabul edilmedi.

Bu davranış bilinçli: Excel'deki eski/referans birim fiyatları kaynağı ve yılı kanıtlanmadan güncel bütçe fiyatı yapılmaz.

## 5. Kod kapsamı ve tasarım sınırları

Önceki kapsam ölçümünde toplam satır kapsamı %47, ana `app.py` kapsamı %44 idi. Bu denetim turunda kapsam yüzdesi yeniden ölçülmedi; test sayısı artışı kapsamın tamamlandığı anlamına gelmez. `app.py` hâlâ 5.000 satırın üzerindedir. Kesik değildir; bakım kolaylığı için veri erişimi, güvenlik, CAD, YFK ve UI modüllerine kontrollü refaktör ayrı iş olarak tutulmalıdır.

## 6. Canlı test bekleyenler

| Alan | Durum | Gereken |
|---|---|---|
| Gerçek Streamlit başlangıcı / browser / staging | BLOKE | Streamlit runtime ve güncel deploy URL |
| OpenAI / Claude / Gemini canlı API | ATLANDI | Gerçek API anahtarları, model erişimi ve açık veri aktarım izni |
| Tavily web fiyat araştırması | ATLANDI | Tavily anahtarı ve canlı API çağrısı |
| ÇŞİDB/YFK canlı senkronizasyonu | ATLANDI | Canlı internet erişimi ve resmî kaynak yanıtı |
| DWG→DXF | BLOKE | ODA converter ve ayrıcalığı kısıtlı izole worker |
| IFC Qto | ATLANDI | `ifcopenshell` ve gerçek Qto içeren IFC kabul modeli |
| ClamAV | ATLANDI | Tarayıcı, güncel imza veritabanı, temiz ve EICAR kabul testi |
| PostgreSQL eşzamanlılık | ATLANDI | Gerçek PostgreSQL ve iki bağımsız oturumla kilit testi |
| Tam yedek/geri yükleme | ATLANDI | DB + dosya varlıkları + kullanıcılar + secrets dışı config restore tatbikatı |

## 7. Üretim öncesi açık ürün işleri

1. KDV dahil/hariç/net tutarları ortak muhasebe esasına göre ayrı saklayan finansal model tamamlanmalı.
2. DWG dönüşüm worker'ı gerçek sandbox/container sınırlarıyla kurulmalı; ODA çağrısı uygulama sürecinde doğrudan güvenli sandbox sayılmaz.
3. KVKK saklama, arşivleme, silme ve imha akışı uygulanmalı.
4. MFA/SSO ve merkezi erişim/rate-limit politikası tamamlanmalı.
5. Dosya varlıkları için kalıcı nesne depolama, yedekleme ve geri yükleme prosedürü canlı ortamda doğrulanmalı.
6. Gerçek browser E2E ve çok kullanıcılı PostgreSQL testi yapılmalı.
7. YFK ve piyasa fiyatlarının güncelliği, tedarikçi teklifiyle ve tarih/teknik uygunluk kanıtıyla karşılaştırılmalı.

## 8. Diğer AI sağlayıcıları hakkında dürüst durum

Uygulama OpenAI, Anthropic Claude, Google Gemini ve OpenAI-uyumlu servisler için sağlayıcı katmanı içerir. Ancak bu denetim sırasında harici sağlayıcıların gerçek API anahtarları mevcut değildi; dolayısıyla Claude/Gemini/OpenAI API çağrıları canlı çalıştırılmadı. Bu rapor GPT-6 destekli inceleme, yerel kod analizi ve otomatik test çıktısına dayanır; harici modellerle ortak canlı inceleme yapılmış gibi sunulmaz.

## 9. Yayın kararı

**V15.5.1 güvenlik düzeltmeli QA adayıdır; üretim nihai sürümü değildir.** Otomatik ve sentetik kabul testleri geçti. Canlı servis, browser, CAD/IFC, antivirüs, PostgreSQL ve tam yedek geri yükleme testleri tamamlanmadan üretime alınması önerilmez.
