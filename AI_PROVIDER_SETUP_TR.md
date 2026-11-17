# V15.5.0 — AI sağlayıcı kurulumu

## Desteklenen bağlantılar

| Sağlayıcı | API | Gerekli Secrets |
|---|---|---|
| OpenAI API (GPT-6) | Responses API | `OPENAI_API_KEY`, `OPENAI_MODEL` |
| Anthropic Claude | Messages API | `ANTHROPIC_API_KEY`, `ANTHROPIC_MODEL` |
| Google Gemini | Interactions API (`/v1/interactions`) | `GEMINI_API_KEY`, `GEMINI_MODEL` |
| OpenAI uyumlu endpoint | Chat Completions | `AI_COMPAT_BASE_URL`, `AI_COMPAT_API_KEY`, `AI_COMPAT_MODEL` |

OpenAI için örnek GPT-6 API model kimliği `gpt-6-astra` olarak önceden yazılmıştır; yalnız hesabınızda erişim varsa kullanın. Diğer model alanlarına sağlayıcı hesabınızda etkin olan tam model kimliğini girin. Örnek veya tahmini model adları kullanmayın. ChatGPT web planı, API anahtarı/ücreti veya model erişimini otomatik sağlamaz.

## Önerilen güvenli başlangıç

```dotenv
AI_PROVIDER=auto
AI_FALLBACK_ENABLED=0
AI_FALLBACK_ORDER=openai,anthropic,gemini,openai_compatible
AI_TIMEOUT_SECONDS=90
AI_MAX_OUTPUT_TOKENS=4096
AI_REVIEW_MAX_PROVIDERS=3
AI_STORE_RESPONSES=0
OPENAI_API_KEY=
OPENAI_MODEL=gpt-6-astra
ANTHROPIC_API_KEY=
ANTHROPIC_MODEL=
GEMINI_API_KEY=
GEMINI_MODEL=
GEMINI_API_VERSION=v1
AI_COMPAT_BASE_URL=
AI_COMPAT_API_KEY=
AI_COMPAT_MODEL=
```

1. Streamlit Cloud / sunucu Settings > Secrets veya sunucu environment içine alanları ekleyin. Anahtarları `app.py`, `.env.example`, git veya ZIP dosyasına yazmayın.
2. Tek sağlayıcıyla başlangıç yapacaksanız yalnız o sağlayıcının key/model alanlarını doldurun.
3. `AI_PROVIDER=auto`, sırada ilk hazır sağlayıcıyı seçer. `AI_FALLBACK_ENABLED=0` ile hata olduğunda başka sağlayıcıya otomatik veri aktarımı olmaz.
4. Yedek sağlayıcıyı açmak isterseniz (`AI_FALLBACK_ENABLED=1`) önce hangi proje verilerinin başka bir şirkete aktarılabileceğini, sözleşme/gizlilik/konum politikalarını ve maliyeti değerlendirip yazılı onay alın. Sağlayıcı hatasında girdi başka bir sağlayıcıya da gönderilebilir.
5. Çoklu AI Uzman Denetimi, ayrı bir onay kutusuyla aynı sınırlı proje özetini hazır seçilen sağlayıcıların tümüne yollar. Her sağlayıcı sonucu ayrı kalır. Sonuçların hiçbirisi otomatik olarak bütçe, WBS, metraj, hakediş, satın alma, ödeme veya sözleşmeye yazılmaz.
6. Model adları, sağlayıcının hesabınızda o tarihte sunduğu model kataloglarından seçilmelidir. Kod model kimliği uydurmaz veya herkese açık en yeni modelin hesabınızda açık olduğunu varsaymaz.

## Sağlayıcı testleri

Paket içindeki `tests/test_ai_gateway.py`, sağlayıcı seçimi, hatalı yapılandırma, fallback, veri formatı dönüştürme, Gemini Interactions endpoint/header ve mock istek/yanıt işleyişini test eder. Gerçek ağ üzerinden erişim, ücretlendirme, rate limit, hesap/model yetkisi ve veri işleme koşulları için her sağlayıcının gerçek API anahtarıyla ayrı kabul testi yapılmalıdır.

## Veri koruma ve kontrol ilkeleri

- API anahtarları arayüzde maskelenir; uygulama kaynak kodunda anahtar bulunmaz.
- OpenAI Responses (`store=false`) ve Gemini Interactions çağrılarında saklama parametresi varsayılan olarak kapalıdır (`AI_STORE_RESPONSES=0`). Bu, sağlayıcının yasal/operasyonel loglarını veya kendi veri politikalarını ortadan kaldırmaz.
- Çoklu sağlayıcı denetimi kullanıcı onayı ister. Yalnız özetlenmiş, sütunları seçilmiş proje verisi gönderilir; tam kaynak ZIP veya orijinal çizimler otomatik aktarılmaz.
- Otomatik fallback varsayılan kapalıdır.
- AI bulguları karar desteğidir. Teknik metraj, statik güvenlik, sözleşme yorumu, fatura, bütçe ve hakediş onayı sorumlu profesyonel kullanıcıda kalır.
