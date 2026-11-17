#!/usr/bin/env python3
"""Pre-deployment checks for Emir Asel Project Control ERP.

Checks declared runtime imports, supported database drivers, configuration
values and local write access. It never prints secret values and does not
claim to validate external integrations end-to-end.
"""
from __future__ import annotations

import importlib.util
import os
import shutil
import sys
import tempfile
import subprocess
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parents[1]
os.chdir(ROOT)
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from security_config import is_placeholder_secret

try:
    from dotenv import load_dotenv
    load_dotenv(ROOT / ".env", override=False)
except Exception:
    pass


try:
    import tomllib
    _secrets_path = ROOT / ".streamlit" / "secrets.toml"
    _secrets = tomllib.loads(_secrets_path.read_text(encoding="utf-8")) if _secrets_path.exists() else {}
except Exception:
    _secrets = {}

def get_setting(name: str, default: str = "") -> str:
    """Mirror app.py precedence: process environment overrides flat Streamlit secrets."""
    value = os.getenv(name)
    if value is not None and value.strip():
        return value.strip()
    secret_value = _secrets.get(name, default) if isinstance(_secrets, dict) else default
    return str(default if secret_value is None else secret_value).strip()

# Modules imported directly by app.py or required by an advertised standard
# feature (official price-source sync / DXF preview) are treated as required.
REQUIRED = {
    "streamlit": "streamlit",
    "pandas": "pandas",
    "numpy": "numpy",
    "plotly": "plotly",
    "sqlalchemy": "sqlalchemy",
    "openpyxl": "openpyxl",
    "ezdxf": "ezdxf",
    "pdfplumber": "pdfplumber",
    "fitz (PyMuPDF)": "fitz",
    "matplotlib (DXF preview)": "matplotlib",
    "requests (official price-source sync)": "requests",
    "bs4 (official price-source parsing)": "bs4",
}
OPTIONAL = {
    "openai (AI features; also requires API key)": "openai",
    "psycopg (PostgreSQL only)": "psycopg",
    "ifcopenshell (IFC reading)": "ifcopenshell",
    "dotenv (.env loading)": "dotenv",
}
errors: list[str] = []
warnings: list[str] = []


def check_module(label: str, module: str, required: bool) -> bool:
    try:
        present = importlib.util.find_spec(module) is not None
    except (ImportError, ValueError, ModuleNotFoundError):
        present = False
    status = "OK" if present else ("MISSING" if required else "OPTIONAL")
    print(f"{status:7} {'required' if required else 'optional'}: {label}")
    if required and not present:
        errors.append(f"Zorunlu bağımlılık eksik: {label}")
    return present


def check_int_env(name: str, default: int, minimum: int, maximum: int) -> int:
    raw = get_setting(name, str(default))
    try:
        value = int(raw)
    except (TypeError, ValueError):
        errors.append(f"{name} tam sayı olmalıdır.")
        return default
    if not minimum <= value <= maximum:
        errors.append(f"{name} değeri {minimum}–{maximum} aralığında olmalıdır.")
        return default
    return value


print(f"Emir Asel ERP preflight | Python {sys.version.split()[0]}")
if sys.version_info < (3, 11):
    errors.append("Python 3.11 veya üzeri gereklidir.")

for label, module in REQUIRED.items():
    check_module(label, module, True)
optional_present = {label: check_module(label, module, False) for label, module in OPTIONAL.items()}

# Validate the backend the application actually supports. SQL statements use
# SQLite/PostgreSQL ON CONFLICT syntax; accepting arbitrary URL schemes here
# would defer a predictable failure until the first database write.
url = get_setting("DATABASE_URL", "")
if not url:
    warnings.append("DATABASE_URL tanımlı değil; uygulama yerel SQLite dosyasına düşer. Bulut/çok kullanıcılı kullanımda kalıcı PostgreSQL önerilir.")
    is_postgres = False
else:
    scheme = urlparse(url).scheme.lower()
    is_sqlite = scheme.startswith("sqlite")
    is_postgres = scheme in {"postgres", "postgresql", "postgresql+psycopg"}
    if not (is_sqlite or is_postgres):
        errors.append("DATABASE_URL yalnızca SQLite veya PostgreSQL (psycopg) kullanabilir.")
    elif is_sqlite:
        warnings.append("DATABASE_URL SQLite gösteriyor; çok kullanıcılı/bulut üretiminde dosya tabanlı veritabanı kalıcılık ve eşzamanlılık riski taşır.")
    else:
        if not optional_present.get("psycopg (PostgreSQL only)", False):
            errors.append("PostgreSQL DATABASE_URL kullanılıyor ancak psycopg sürücüsü kurulu değil.")
        print("OK      DATABASE_URL PostgreSQL olarak tanımlı (değer gizlendi).")

# A short or copied example bootstrap secret must not be used on first run.
token = get_setting("EA_BOOTSTRAP_TOKEN", "")
if not token:
    warnings.append("EA_BOOTSTRAP_TOKEN yok. Yeni veritabanında ilk yönetici hesabı oluşturulamaz.")
elif is_placeholder_secret(token):
    errors.append("EA_BOOTSTRAP_TOKEN en az 32 karakterlik, örnekten farklı rastgele bir değer olmalıdır.")
else:
    print("OK      EA_BOOTSTRAP_TOKEN mevcut; değeri gizlendi.")

ai_provider_specs = {
    "OpenAI": ("OPENAI_API_KEY", "OPENAI_MODEL"),
    "Anthropic Claude": ("ANTHROPIC_API_KEY", "ANTHROPIC_MODEL"),
    "Google Gemini": ("GEMINI_API_KEY", "GEMINI_MODEL"),
    "OpenAI uyumlu servis": ("AI_COMPAT_API_KEY", "AI_COMPAT_MODEL"),
}
ready_ai = []
partial_ai = []
for provider_name, (key_name, model_name) in ai_provider_specs.items():
    has_key = bool(get_setting(key_name, ""))
    has_model = bool(get_setting(model_name, ""))
    if has_key and has_model:
        if provider_name == "OpenAI" and not optional_present.get("openai (AI features; also requires API key)", False):
            errors.append("OpenAI API anahtarı/modeli tanımlı ancak openai SDK paketi kurulu değil.")
        elif provider_name == "OpenAI uyumlu servis" and not get_setting("AI_COMPAT_BASE_URL", ""):
            partial_ai.append(provider_name)
        else:
            ready_ai.append(provider_name)
    elif has_key or has_model:
        partial_ai.append(provider_name)
if not ready_ai:
    warnings.append("Hiçbir AI sağlayıcısında hem API anahtarı hem model adı yok; AI özellikleri devre dışı kalır.")
else:
    print("OK      AI sağlayıcı yapılandırması: " + ", ".join(ready_ai) + " (anahtar değerleri gizlendi). Canlı API çağrısı ayrıca doğrulanmalıdır.")
if partial_ai:
    warnings.append("AI sağlayıcılarında eksik key/model çifti var: " + ", ".join(partial_ai) + ".")
if not get_setting("TAVILY_API_KEY", ""):
    warnings.append("TAVILY_API_KEY yok; internetten tedarikçi/piyasa fiyatı arama devre dışı kalır. Fiyat uydurulmaz; yerel kütüphane/YFK kaynakları kullanılabilir.")
else:
    print("OK      TAVILY_API_KEY mevcut; web fiyat arama servisi için ayarlanmış (değer gizlendi). Gerçek API çağrısı ayrıca test edilmelidir.")

max_upload_mb = check_int_env("EA_MAX_UPLOAD_MB", 100, 1, 1000)
server_upload_mb = 100
try:
    import tomllib
    config = tomllib.loads((ROOT / ".streamlit" / "config.toml").read_text(encoding="utf-8"))
    server_upload_mb = int(config.get("server", {}).get("maxUploadSize", 100))
except Exception:
    warnings.append(".streamlit/config.toml maxUploadSize ayarı okunamadı; Streamlit varsayılanı 200 MB olabilir.")
if max_upload_mb > server_upload_mb:
    warnings.append(f"EA_MAX_UPLOAD_MB={max_upload_mb}, Streamlit maxUploadSize={server_upload_mb}; uygulama sınırını Streamlit ayarı da karşılamalı.")

check_int_env("EA_MAX_OOXML_ENTRIES", 10000, 1, 100000)
check_int_env("EA_MAX_OOXML_UNCOMPRESSED_MB", 256, 1, 4096)
ratio_raw = get_setting("EA_MAX_OOXML_RATIO", "150")
try:
    ratio = float(ratio_raw)
    if not (1 <= ratio <= 10000):
        raise ValueError
except (TypeError, ValueError):
    errors.append("EA_MAX_OOXML_RATIO 1–10000 aralığında bir sayı olmalıdır.")

if get_setting("ENABLE_ODA_CONVERSION", "0") == "1":
    exe = get_setting("ODA_FILE_CONVERTER", "")
    if not exe or not Path(exe).is_file():
        errors.append("ENABLE_ODA_CONVERSION=1 ancak ODA_FILE_CONVERTER geçerli bir dosyayı göstermiyor.")
    else:
        print("OK      ODA dönüştürücü yolu mevcut; gerçek DWG kabul testi yine gerekir.")
else:
    warnings.append("DWG dönüştürme kapalı; DXF analizi kullanılabilir, DWG için izole ODA worker yapılandırılmalı.")

clam = get_setting("CLAMSCAN_PATH", "")
if clam:
    if Path(clam).is_file():
        try:
            probe = subprocess.run([clam, "--version"], capture_output=True, text=True, timeout=10, shell=False)
            if probe.returncode != 0 or not (probe.stdout or probe.stderr).strip():
                errors.append("ClamAV yolu mevcut ancak --version kontrolü başarısız.")
            else:
                print("OK      ClamAV executable çalıştırılabiliyor; imza veritabanı/gerçek dosya taraması ayrıca doğrulanmalı.")
        except Exception as exc:
            errors.append(f"ClamAV çalıştırılamadı: {type(exc).__name__}: {exc}")
    else:
        errors.append("CLAMSCAN_PATH geçersiz.")
else:
    warnings.append("Antivirüs taraması yapılandırılmamış; dosyalar taranmadan kabul edilir. Üretimde ClamAV/eşdeğeri zorunludur.")

if get_setting("EA_ENV", "").lower() == "production":
    if not is_postgres:
        errors.append("EA_ENV=production için kalıcı PostgreSQL DATABASE_URL zorunludur.")
    if is_placeholder_secret(token):
        errors.append("EA_ENV=production için örnek olmayan, en az 32 karakterlik rastgele EA_BOOTSTRAP_TOKEN zorunludur.")
    if not clam:
        errors.append("EA_ENV=production için CLAMSCAN_PATH antivirüs tarayıcısı zorunludur.")

# Check that the application can create its local data/upload directories.
try:
    data_dir = ROOT / "data"
    upload_dir = data_dir / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix=".ea_preflight_", dir=upload_dir, delete=True):
        pass
    print("OK      Yerel veri/dosya alanına yazılabiliyor.")
except Exception as exc:
    errors.append(f"Yerel data/uploads dizinine yazılamıyor: {type(exc).__name__}: {exc}")

print("\nWARNINGS")
for item in warnings:
    print(f"- {item}")
print("\nERRORS")
for item in errors:
    print(f"- {item}")
if errors:
    raise SystemExit(1)
print("\nPreflight tamamlandı. Bu kontrol canlı entegrasyon veya üretim kabul testi yerine geçmez.")
