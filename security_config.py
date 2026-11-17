"""Shared validation for bootstrap secrets and deployment configuration."""
from __future__ import annotations
import re
import html

_PLACEHOLDER_VALUES = {
    "UZUN_RASTGELE_ILK_KURULUM_ANAHTARI",
    "BURAYA_32_KARAKTERDEN_UZUN_RASTGELE_TOKEN",
    "EN_AZ_32_KARAKTER_RASTGELE_GIZLI_DEGER",
    "REPLACE_WITH_RANDOM_SECRET_MIN_32_CHARS",
    "CHANGE_ME",
    "YOUR_SECRET_HERE",
}


def is_placeholder_secret(value: object, minimum_length: int = 32) -> bool:
    """Return True for short or documented example secrets, even if long enough."""
    raw = str(value or "").strip()
    if len(raw) < minimum_length:
        return True
    normalized = re.sub(r"[^A-Z0-9]+", "_", raw.upper()).strip("_")
    if normalized in _PLACEHOLDER_VALUES:
        return True
    return normalized.startswith(("BURAYA_", "REPLACE_", "CHANGE_ME_", "YOUR_SECRET_", "EXAMPLE_"))


def escape_html_text(value: object) -> str:
    """Escape untrusted user-controlled text before inserting it into HTML."""
    return html.escape(str(value or ""), quote=True)
