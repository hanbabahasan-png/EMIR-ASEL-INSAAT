# GitHub'a yükleme ve Streamlit Cloud dağıtım rehberi

## 1. GitHub deposuna yükleme

1. `https://github.com/hanbabahasan-png/EMIR-ASEL-INSAAT` deposunu açın.
2. `main` dalında mevcut dosyaların bir kopyasını/geri dönüş commit'ini saklayın.
3. Bu paketin içindeki klasörleri değil, klasörün **içindeki tüm dosya ve klasörleri** deponun kök dizinine yükleyin. `app.py` doğrudan kökte olmalıdır.
4. GitHub web arayüzünde `Add file > Upload files` ile dosyaları yükleyin veya Git kullanarak commit/push yapın.
5. `.streamlit/secrets.toml.example` dosyasını GitHub'a yükleyebilirsiniz; gerçek `secrets.toml` dosyasını YÜKLEMEYİN.

## 2. Streamlit Cloud

1. Streamlit Community Cloud'da mevcut uygulamanın ayarlarını açın.
2. Repository: `hanbabahasan-png/EMIR-ASEL-INSAAT`
3. Branch: `main`
4. Main file path: `app.py`
5. Advanced settings / Secrets bölümüne gerçek sırları güvenli biçimde girin.
6. Uygulamayı yeniden başlatın ve logları kontrol edin.

## 3. Gerekli üretim Secrets

`.streamlit/secrets.toml.example` dosyasını örnek alın. Gerçek Neon PostgreSQL bağlantı adresini `DATABASE_URL` alanına girin. Uygulama için Neon pooled bağlantı adresi kullanılabilir; migration/backup işlemleri için doğrudan bağlantı ayrı tutulmalıdır.

En azından:

- `EA_ENV = "production"`
- `EA_BOOTSTRAP_TOKEN = "32+ karakter rastgele gizli değer"`
- `EA_BASE_CURRENCY = "TRY"`
- `DATABASE_URL = "postgresql+psycopg://..."`

AI özellikleri için ilgili sağlayıcının API anahtarı ve hesabınızda etkin model kimliği gerekir. İnternet fiyat araştırması için `TAVILY_API_KEY` gerekir. Bu anahtarları Python koduna veya GitHub'a koymayın.

## 4. Güvenlik uyarıları

- `.streamlit/secrets.toml`, `.env`, veritabanı dosyaları ve yüklenen proje belgeleri Git'e gönderilmemelidir.
- Neon veritabanı yedeği ve dosya deposu yedeği alınmadan üretim verisiyle işlem yapmayın.
- İlk dağıtımı mümkünse staging/test veritabanında yapın.
- Bu paket gerçek Neon kimlik bilgilerini veya Streamlit hesabına erişimi içermez; bu nedenle canlı bağlantının başarılı olduğu iddia edilmez.

## 5. Dağıtım sonrası kabul

1. Uygulama açılıyor ve Sistem Bilgileri ekranında `15.5.1` gösteriyor.
2. Veritabanı türü PostgreSQL olarak görünüyor.
3. Test projesi Excel dosyasını içe aktarın.
4. Metraj, fiyat kaynağı, bütçe, onay, gerçekleşen, hakediş ve rapor akışlarını kontrol edin.
5. Test kaydı oluşturup uygulamayı yeniden başlatın; kaydın kaldığını doğrulayın.
6. Üretim verisi kullanmadan önce yedekten geri yükleme provası yapın.
