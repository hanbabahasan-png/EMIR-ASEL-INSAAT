# Ünzüle Hanbaba — V15.4.1 gerçek proje paketi ön kabul denetimi

Tarih: 16.11.2026

## Kaynak paketi envanteri
- 4 DWG: mimari, statik, tesisat, zayıf akım. Dosya başlığı AC1032 (AutoCAD 2018).
- PDF: 162 sayfa statik hesap raporu, 8 sayfa statik proje, 122 sayfa zemin etüdü, 34 sayfa İZODER, 3 sayfa klima hesabı, 1 sayfa Enerjisa.
- Excel: FINAL4 ana giderler, FINAL2/FINAL3 doldurulmuş metraj, hiyerarşik bütçe v3, çapraz metraj v2.
- Kaynak ZIP CRC kontrolü başarılı; 17 arşiv girdisi bulundu.

## Kodun gerçek proje için mevcut sınırları
- PDF AI takeoff tek çalıştırmada en fazla 24 sayfa işler. 162 sayfalık rapor pafta/konu gruplarına ayrılmalıdır.
- DWG dönüşümü varsayılan kapalıdır; ODA File Converter ayrı, kısıtlı worker üzerinde yapılandırılmadan bu dört DWG'den otomatik CAD metrajı alınamaz.
- Canlı AI ve internet fiyat araştırması için OPENAI_API_KEY / OPENAI_MODEL / TAVILY_API_KEY gerekir.
- Bu çalışma ortamında Streamlit paketi yok; production preflight bu nedenle başarısızdır. PostgreSQL, ClamAV, ODA worker ve canlı API entegrasyonu da doğrulanmamıştır.

## Proje dosyalarında görülen kontrol başlıkları
1. `Villa_Insaat_Hiyerarsik_Butce_v3.xlsx` içindeki 160 WBS satırının 72'si AÇIK, 43'ü HESAP, 33'ü TAHMİN, 12'si DOĞRULANDI/DOGRULANDI olarak işaretlidir. Bazı `DOGRULANDI` satırlarının açıklaması "Tipik oran" olduğundan statü ile kanıt açıklaması çelişebilir; statik pafta ve metraj cetveliyle yeniden doğrulanmalıdır.
2. Çapraz metraj dosyasında 3.673 t, 1.73425 t, 2.35197 t, 2.103 t ve 1.66234 t gibi donatı grup değerleri vardır. Ayrı tabloda 2.35197 t kiriş grubu iki başlıkta tekrar görünür; ayrıca 0.47539 t alt grup satırı vardır. Bunlar bina toplamı olarak körlemesine toplanmamalıdır. Kaynak statik PDF pafta 4 ve 5 ile grup/alt grup ilişkisi mühendis tarafından teyit edilmelidir.
3. Çapraz metraj tablosundaki 30.500 TL/ton demir ve 3.700 TL/m³ beton gibi fiyatlarda açık fiyat yılı/teklif kaynağı yoktur. Bunlar güncel piyasa fiyatı kabul edilmemelidir. Gerçek `Metraj` sayfası V15.4.1 içe aktarım motorunda denendi: 13 satır okundu; 12 fiyatın tamamı kaynak/yıl eksikliği nedeniyle bütçe BF alanında boş bırakıldı, tutarlar açıklamada referans olarak korundu. V15.4.1 bu tür fiyatların otomatik olarak 2026 yılına ve dosya adına bağlanmasını engeller.
4. Statik projede 199,7 m² alan bilgisi bulunurken klima hesabı kapağında 251,20 m² alan yazmaktadır. Alan tanımları (brüt/net, toplam/ısıtılan alan, kat dahil/hariç) eşleştirilmeden tek bir alan olarak kullanılmamalıdır.
5. Çatı iskelesi/çapraz çelik ağırlığı ile kalıcı çatı taşıyıcı çeliği ayrı kapsamdır; aynı WBS kalemine otomatik eşlenmemelidir.

## Test ve kabul durumu
- Gerçek proje Excel kuru çalıştırma: `Villa_Insaat_Proje_Capraz_Metraj_v2.xlsx` / `Metraj` sayfası, 13 satır okundu; WBS eşlemesi kullanıcı onayı olmadan yapılmadı; 12 fiyat kaynağı/yılı eksik olduğu için fiyat olarak aktarılmadı.
- Kaynak V15.4.0 paketi: `pytest -q` 67 passed.
- UI shell: 34/34 sayfa; mock Streamlit kabuğu, gerçek tarayıcı değildir.
- E2E kabul: 38 passed, 0 failed, 2 skipped (gerçek IFC ve canlı OpenAI).
- V15.4.1 fiyat kaynağı düzeltmesi için ek 6 test bu teslim paketinde ayrıca çalıştırılacaktır.
- Bu rapor gerçek projedeki DWG geometrilerinin ölçüldüğü veya canlı AI metrajı üretildiği anlamına gelmez.
