# EMİR ASEL İNŞAAT — V15.5.1 Canlı Kabul Durumu

**Durum: CANLI ÜRETİM KABULÜ BEKLEMEDE.** Otomatik ve test proje kontrolleri başarılı; gerçek canlı servisler ve kurumsal altyapı kabul edilmedi.

## Otomatik testler
- Python derleme: GEÇTİ
- pytest: **100 geçti, 0 başarısız** (+14 subtest)
- UI kabuk testi: **34/34** (mock Streamlit)
- Sentetik proje E2E: **38 geçti, 0 başarısız, 2 atlandı**
- Toplam ölçülen satır kapsamı: **%47**

Atlanan testler gerçek IFC dosya entegrasyonu ve canlı AI API'dir.

## Gerçek Ünzüle Hanbaba kaynakları
- 6 PDF, 5 Excel ve 4 DWG dosyası dosya formatı/imzası üzerinden kontrol edildi; Excel dosyaları açıldı, PDF sayfa sayısı kontrol edildi, DWG sürüm imzası AC1032 olarak tanındı.
- Dört DWG dosyasında ODA File Converter olmadığı için geometri/metraj hesaplanmadı.
- Gerçek `Villa_Insaat_Proje_Capraz_Metraj_v2.xlsx` / `Metraj` sayfası: 13 aday, 8 taslak aktarım adayı, kaynak/yıl doğrulaması olmadığından bütçeye kabul edilen fiyat 0.
- ClamAV yok; dosyalar antivirüs taramasından geçmedi.

## Canlı kabul engelleri
1. Streamlit bu test ortamında kurulu değil; gerçek uygulama başlangıç ve browser testleri yapılamadı. Pip kurulumu DNS/ağ erişimi yokluğu nedeniyle başarısız oldu.
2. OpenAI, Claude, Gemini ve Tavily için canlı API anahtarı/erişim yok.
3. YFK resmî senkronizasyonu canlı HTTP üzerinden denenmedi.
4. ODA DWG converter, `ifcopenshell`, ClamAV, PostgreSQL ve `psycopg` hazır değil.
5. PostgreSQL iki kullanıcı eşzamanlı yazma ve tam DB+dosya yedek/geri yükleme tatbikatı yapılmadı.
6. Gerçekleşen faturalarda KDV dahil/hariç ayrıştırması otomatik değil; raporlarda aynı maliyet esası kullanılmalı.
7. KVKK saklama/silme politikası ile MFA/SSO tamamlanmamış.

Üretime geçiş için Streamlit staging URL'si ve yukarıdaki servis/anahtarların yetkili ortamda yapılandırılması gerekir.


## Bu denetim turunda eklenen güvenlik düzeltmeleri
- İlk yönetici bootstrap placeholder anahtarları uygulama ve preflight seviyesinde reddedilir.
- Yönetici panelinde kullanıcı kontrollü proje adı HTML escape edilir.
- Son otomatik test: 105 pytest geçti; 34/34 UI kabuk; 38 E2E geçti, 2 harici entegrasyon atlandı.
