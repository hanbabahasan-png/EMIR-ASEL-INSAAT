# V15.4.1 — Excel fiyat kaynağı / yıl güvenlik düzeltmesi

## Kritik düzeltme
V15.4.0 Excel metraj aktarımında, fiyat yılı boşsa bugünün yılını; fiyat kaynağı boşsa Excel dosya adını kaynak gibi atayabiliyordu. Bu davranış, eski/teyitsiz birim fiyatların güncel ve kaynaklıymış gibi görünmesine yol açabilirdi.

V15.4.1 ile:
- Excel'den gelen fiyat, açık fiyat yılı ve açık fiyat kaynağı birlikte bulunmadıkça bütçe birim fiyatına aktarılmaz.
- İçe aktarılan fiyat 2 yıl veya daha eskiyse bütçe fiyatı olarak kabul edilmez; yeni teklif doğrulaması bekler.
- Eski referans tutar silinmez; satırın açıklamasında referans olarak korunur.
- Fiyatı eksik satır metraj taslağı olarak aktarılabilir; ancak fiyat boş kalır.
- Gerçek kaynak ve fiyat yılı girilmiş, güncellik aralığındaki tutarlar korunur.

## Korunan modüller
V15.4.0 modülleri ve mevcut veritabanı şeması korunmuştur. Bu düzeltme yeni tablo/kolon eklemez.

## Test
`tests/test_price_provenance.py`: 6 yeni regresyon testi.
