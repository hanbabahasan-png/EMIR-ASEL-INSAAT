from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from security_config import is_placeholder_secret, escape_html_text


def test_short_or_empty_bootstrap_secret_is_rejected():
    assert is_placeholder_secret("")
    assert is_placeholder_secret("short-secret")


def test_all_documented_bootstrap_placeholders_are_rejected():
    assert is_placeholder_secret("UZUN_RASTGELE_ILK_KURULUM_ANAHTARI")
    assert is_placeholder_secret("BURAYA_32_KARAKTERDEN_UZUN_RASTGELE_TOKEN")
    assert is_placeholder_secret("EN_AZ_32_KARAKTER_RASTGELE_GIZLI_DEGER")
    assert is_placeholder_secret("REPLACE_WITH_RANDOM_SECRET_MIN_32_CHARS")


def test_generated_high_entropy_style_secret_is_accepted():
    assert not is_placeholder_secret("v7!kQ2#pL9@xR4$zN8%mT1^bH6&cF3*wS5!qJ0")


def test_secrets_example_does_not_ship_a_usable_bootstrap_value():
    content = (ROOT / ".streamlit" / "secrets.toml.example").read_text(encoding="utf-8")
    line = next(x for x in content.splitlines() if x.strip().startswith("EA_BOOTSTRAP_TOKEN"))
    value = line.split("=", 1)[1].split("#", 1)[0].strip().strip('"')
    assert is_placeholder_secret(value)


def test_user_controlled_project_name_is_html_escaped():
    malicious = '<img src=x onerror="alert(1)">'
    escaped = escape_html_text(malicious)
    assert escaped == "&lt;img src=x onerror=&quot;alert(1)&quot;&gt;"
    assert "<img" not in escaped
