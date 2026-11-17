# Emir Asel İnşaat — V15.5.1 Sürüm Notları

## Düzeltmeler
- Bootstrap token placeholder açığı kapatıldı; uygulama ve preflight aynı doğrulamayı kullanır.
- Yönetici dashboard proje adı HTML enjeksiyonuna karşı escape edilir.
- E2E test geçici paketine ortak güvenlik modülü eklendi.
- Onaylı iş/finans kayıtlarında merkezi değişiklik koruması: maddi alan değişikliğinde yeniden onay; onaylı kayıt silme engeli.
- Metraj adayları için kaynak sayfa, yayın tarihi ve kontrol tarihi alanları; `DATE_FIELDS` normalizasyonu.
- Piyasa tekliflerinde kaynak SHA256, kaynak dosya ve araştırma sorgusu izleri.
- Excel metraj importunda `Kalem` kolonunun otomatik tanınması; genel `Kaynak Durumu` alanının fiyat kaynağı olarak yanlış kabul edilmemesi.
- Stok maliyet, ağırlıklı ilerleme ve veri kalite testleri.

## Test özeti
- Python derleme: geçti.
- pytest: 105 passed, 14 subtests passed, 0 failed.
- UI shell: 34/34 passed (mock Streamlit).
- Sentetik E2E: 38 passed, 0 failed, 2 skipped (IFC live file + AI live API).
- Ölçülen satır kapsamı: toplam %47 (app.py %44, ai_gateway.py %63, operations_monitor.py %88).

## Üretim kabulü
**Henüz verilmedi.** Streamlit gerçek browser/staging, canlı AI, Tavily, YFK, DWG ODA dönüşümü, IFC, ClamAV, PostgreSQL eşzamanlılık ve tam yedek/geri yükleme tatbikatı ortam gereksinimleri nedeniyle tamamlanmadı. KDV dahil/hariç maliyet ayrıştırması ile KVKK saklama/silme ve MFA/SSO da açık ürün işleri olarak kayıtlıdır.
