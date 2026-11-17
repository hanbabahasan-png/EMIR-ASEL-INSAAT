# -*- coding: utf-8 -*-
"""Provider-neutral AI gateway for Emir Asel ERP.

Supports OpenAI Responses, Anthropic Messages, Google Gemini generateContent and
OpenAI-compatible Chat Completions endpoints. No provider key is stored here.
"""
from __future__ import annotations

import base64
import re
import time
from dataclasses import dataclass
from typing import Any

import requests


PROVIDER_LABELS = {
    "openai": "OpenAI / ChatGPT API",
    "anthropic": "Anthropic Claude",
    "gemini": "Google Gemini",
    "openai_compatible": "OpenAI uyumlu servis",
}
PROVIDER_KEY_NAMES = {
    "openai": "OPENAI_API_KEY",
    "anthropic": "ANTHROPIC_API_KEY",
    "gemini": "GEMINI_API_KEY",
    "openai_compatible": "AI_COMPAT_API_KEY",
}
PROVIDER_MODEL_NAMES = {
    "openai": "OPENAI_MODEL",
    "anthropic": "ANTHROPIC_MODEL",
    "gemini": "GEMINI_MODEL",
    "openai_compatible": "AI_COMPAT_MODEL",
}
DEFAULT_ORDER = ("openai", "anthropic", "gemini", "openai_compatible")


@dataclass
class AIResult:
    output_text: str
    provider: str
    model: str
    fallback_used: bool = False


class AIRouterError(RuntimeError):
    pass


def _truthy(value: Any) -> bool:
    return str(value or "").strip().casefold() in {"1", "true", "yes", "evet", "on"}


def _safe_int(config: dict[str, Any], key: str, default: int, minimum: int, maximum: int) -> int:
    try:
        value = int(str(config.get(key) or default).strip())
    except (TypeError, ValueError):
        value = default
    return max(minimum, min(maximum, value))


def configured_providers(config: dict[str, Any]) -> list[str]:
    order_raw = str(config.get("AI_FALLBACK_ORDER") or ",".join(DEFAULT_ORDER))
    order = [x.strip().lower() for x in re.split(r"[,; ]+", order_raw) if x.strip()]
    order = [x for x in order if x in PROVIDER_LABELS]
    order += [x for x in DEFAULT_ORDER if x not in order]
    return [p for p in order if str(config.get(PROVIDER_KEY_NAMES[p]) or "").strip()
            and str(config.get(PROVIDER_MODEL_NAMES[p]) or "").strip()
            and (p != "openai_compatible" or str(config.get("AI_COMPAT_BASE_URL") or "").strip())]


def provider_status(config: dict[str, Any]) -> list[dict[str, str]]:
    ready = set(configured_providers(config))
    rows = []
    for provider in DEFAULT_ORDER:
        key_name = PROVIDER_KEY_NAMES[provider]
        model_name = PROVIDER_MODEL_NAMES[provider]
        has_key = bool(str(config.get(key_name) or "").strip())
        model = str(config.get(model_name) or "").strip()
        if not has_key:
            state = "API anahtarı eksik"
        elif not model:
            state = "Model adı eksik"
        elif provider == "openai_compatible" and not str(config.get("AI_COMPAT_BASE_URL") or "").strip():
            state = "Base URL eksik"
        elif provider in ready:
            state = "Hazır"
        else:
            state = "Yapılandırma eksik"
        rows.append({"Sağlayıcı": PROVIDER_LABELS[provider], "Durum": state,
                     "Model": model or "—", "Anahtar": "Tanımlı" if has_key else "Tanımlı değil"})
    return rows


def _input_messages(input_data: Any) -> list[dict[str, Any]]:
    if isinstance(input_data, str):
        return [{"role": "user", "content": [{"type": "input_text", "text": input_data}]}]
    if not isinstance(input_data, list):
        raise TypeError("AI girdisi metin veya mesaj listesi olmalıdır.")
    messages: list[dict[str, Any]] = []
    for item in input_data:
        if not isinstance(item, dict):
            continue
        role = str(item.get("role") or "user").lower()
        if role not in {"user", "assistant", "system", "developer"}:
            role = "user"
        content = item.get("content", "")
        if isinstance(content, str):
            parts = [{"type": "input_text", "text": content}]
        elif isinstance(content, list):
            parts = []
            for part in content:
                if not isinstance(part, dict):
                    continue
                ptype = part.get("type")
                if ptype in {"input_text", "text"}:
                    parts.append({"type": "input_text", "text": str(part.get("text") or "")})
                elif ptype in {"input_image", "image_url"}:
                    image_url = part.get("image_url")
                    if isinstance(image_url, dict):
                        image_url = image_url.get("url")
                    if image_url:
                        parts.append({"type": "input_image", "image_url": str(image_url)})
        else:
            parts = [{"type": "input_text", "text": str(content)}]
        if parts:
            messages.append({"role": role, "content": parts})
    return messages


def _text_from_parts(parts: list[dict[str, Any]]) -> str:
    return "\n".join(str(p.get("text") or "") for p in parts if p.get("type") == "input_text").strip()


def _image_data_uri(uri: str) -> tuple[str, str]:
    match = re.fullmatch(r"data:(image/(?:png|jpeg|webp|gif));base64,([A-Za-z0-9+/=\r\n]+)", uri.strip())
    if not match:
        raise ValueError("Bu sağlayıcıya yalnızca base64 data URI görsel girdisi aktarılabilir.")
    return match.group(1), re.sub(r"\s+", "", match.group(2))


def _to_anthropic(messages: list[dict[str, Any]]) -> tuple[str, list[dict[str, Any]]]:
    system_parts: list[str] = []
    converted: list[dict[str, Any]] = []
    for msg in messages:
        role = msg["role"]
        blocks = []
        for part in msg["content"]:
            if part["type"] == "input_text":
                if part.get("text"):
                    blocks.append({"type": "text", "text": str(part["text"])})
            elif part["type"] == "input_image":
                mime, encoded = _image_data_uri(str(part.get("image_url") or ""))
                blocks.append({"type": "image", "source": {"type": "base64", "media_type": mime, "data": encoded}})
        if not blocks:
            continue
        if role in {"system", "developer"}:
            system_parts.extend(b.get("text", "") for b in blocks if b["type"] == "text")
        else:
            converted.append({"role": "assistant" if role == "assistant" else "user", "content": blocks})
    return "\n\n".join(system_parts), converted


def _to_gemini_interactions(messages: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Convert the app's text/image messages to the current Gemini Interactions content format."""
    content: list[dict[str, Any]] = []
    for msg in messages:
        role = msg["role"]
        for part in msg["content"]:
            if part["type"] == "input_text":
                text = str(part.get("text") or "")
                if text:
                    if role == "assistant":
                        text = "Önceki asistan yanıtı: " + text
                    elif role in {"system", "developer"}:
                        text = "Bağlam notu: " + text
                    content.append({"type": "text", "text": text})
            elif part["type"] == "input_image":
                mime, encoded = _image_data_uri(str(part.get("image_url") or ""))
                content.append({"type": "image", "data": encoded, "mime_type": mime})
    return content


def _to_chat_messages(instructions: str, messages: list[dict[str, Any]]) -> list[dict[str, Any]]:
    out = []
    if instructions:
        out.append({"role": "system", "content": instructions})
    for msg in messages:
        role = msg["role"] if msg["role"] in {"user", "assistant", "system"} else "user"
        parts = []
        for part in msg["content"]:
            if part["type"] == "input_text":
                parts.append({"type": "text", "text": str(part.get("text") or "")})
            elif part["type"] == "input_image":
                parts.append({"type": "image_url", "image_url": {"url": str(part.get("image_url") or "")}})
        if all(x.get("type") == "text" for x in parts):
            content = "\n".join(x.get("text", "") for x in parts)
        else:
            content = parts
        out.append({"role": role, "content": content})
    return out


def _http_post(url: str, headers: dict[str, str], body: dict[str, Any], timeout: int) -> dict[str, Any]:
    last_error = None
    for attempt in range(2):
        try:
            response = requests.post(url, headers=headers, json=body, timeout=(10, timeout))
            if response.status_code == 429 or response.status_code >= 500:
                response.raise_for_status()
            response.raise_for_status()
            data = response.json()
            if not isinstance(data, dict):
                raise AIRouterError("AI sağlayıcısı beklenmeyen yanıt biçimi döndürdü.")
            return data
        except Exception as exc:  # network/provider exceptions are intentionally sanitized
            last_error = exc
            if attempt == 0 and (not isinstance(exc, requests.HTTPError) or
                                 (exc.response is not None and exc.response.status_code >= 500)):
                time.sleep(0.4)
                continue
            break
    if isinstance(last_error, requests.HTTPError) and last_error.response is not None:
        status = last_error.response.status_code
        raise AIRouterError(f"AI sağlayıcısı HTTP {status} hatası döndürdü.") from last_error
    raise AIRouterError(f"AI sağlayıcısına erişilemedi: {type(last_error).__name__ if last_error else 'bilinmeyen hata'}") from last_error


def _call_openai(config: dict[str, Any], model: str, instructions: str, input_data: Any, compatible: bool = False) -> str:
    try:
        from openai import OpenAI
    except ImportError as exc:
        raise AIRouterError("OpenAI SDK bulunamadı; requirements.txt kurulmalıdır.") from exc
    api_key = str((config.get("AI_COMPAT_API_KEY") if compatible else config.get("OPENAI_API_KEY")) or "").strip()
    if compatible and not api_key:
        api_key = str(config.get("OPENAI_API_KEY") or "").strip()
    base_url = str(config.get("AI_COMPAT_BASE_URL") or "").strip() if compatible else ""
    timeout = _safe_int(config, "AI_TIMEOUT_SECONDS", 90, 10, 300)
    kwargs = {"api_key": api_key, "timeout": timeout, "max_retries": 0}
    if base_url:
        kwargs["base_url"] = base_url
    client = OpenAI(**kwargs)
    if compatible:
        messages = _to_chat_messages(instructions, _input_messages(input_data))
        response = client.chat.completions.create(model=model, messages=messages)
        return str(response.choices[0].message.content or "")
    response = client.responses.create(model=model, instructions=instructions or None, input=input_data,
                                        store=_truthy(config.get("AI_STORE_RESPONSES", "0")))
    output = getattr(response, "output_text", None)
    if output is None:
        raise AIRouterError("OpenAI yanıtında output_text bulunamadı.")
    return str(output)


def _call_anthropic(config: dict[str, Any], model: str, instructions: str, input_data: Any) -> str:
    key = str(config.get("ANTHROPIC_API_KEY") or "").strip()
    system_from_input, messages = _to_anthropic(_input_messages(input_data))
    system_text = "\n\n".join(x for x in (instructions, system_from_input) if x)
    if not messages:
        messages = [{"role": "user", "content": "Lütfen verilen talimatı uygula."}]
    version = str(config.get("ANTHROPIC_VERSION") or "2023-06-01")
    body = {"model": model, "max_tokens": _safe_int(config, "AI_MAX_OUTPUT_TOKENS", 4096, 256, 65536), "messages": messages}
    if system_text:
        body["system"] = system_text
    data = _http_post("https://api.anthropic.com/v1/messages",
                      {"x-api-key": key, "anthropic-version": version, "content-type": "application/json"},
                      body, _safe_int(config, "AI_TIMEOUT_SECONDS", 90, 10, 300))
    return "".join(block.get("text", "") for block in data.get("content", [])
                   if isinstance(block, dict) and block.get("type") == "text").strip()


def _call_gemini(config: dict[str, Any], model: str, instructions: str, input_data: Any) -> str:
    """Use Google's current Interactions API (stable v1) rather than legacy generateContent."""
    key = str(config.get("GEMINI_API_KEY") or "").strip()
    messages = _input_messages(input_data)
    system_parts = [str(instructions).strip()] if instructions else []
    for msg in messages:
        if msg["role"] in {"system", "developer"}:
            system_parts.extend(str(p.get("text") or "") for p in msg["content"] if p["type"] == "input_text")
    contents = _to_gemini_interactions([m for m in messages if m["role"] not in {"system", "developer"}])
    if not contents:
        contents = [{"type": "text", "text": "Lütfen verilen talimatı uygula."}]
    body: dict[str, Any] = {
        "model": model,
        "input": contents,
        "store": _truthy(config.get("AI_STORE_RESPONSES", "0")),
        "generation_config": {"max_output_tokens": _safe_int(config, "AI_MAX_OUTPUT_TOKENS", 4096, 256, 65536)},
    }
    system_text = "\n\n".join(x for x in system_parts if x)
    if system_text:
        body["system_instruction"] = system_text
    version = str(config.get("GEMINI_API_VERSION") or "v1").strip().strip("/")
    if version not in {"v1", "v1beta"}:
        version = "v1"
    url = f"https://generativelanguage.googleapis.com/{version}/interactions"
    data = _http_post(url, {"x-goog-api-key": key, "content-type": "application/json"}, body,
                      _safe_int(config, "AI_TIMEOUT_SECONDS", 90, 10, 300))
    output = data.get("output_text")
    if output:
        return str(output).strip()
    texts = []
    for step in data.get("steps", []) or []:
        for block in step.get("content", []) if isinstance(step, dict) else []:
            if isinstance(block, dict) and block.get("type") == "text":
                texts.append(str(block.get("text") or ""))
    return "\n".join(texts).strip()


def call_provider(config: dict[str, Any], provider: str, instructions: str, input_data: Any) -> AIResult:
    if provider not in PROVIDER_LABELS:
        raise AIRouterError("AI sağlayıcısı tanınmıyor.")
    model = str(config.get(PROVIDER_MODEL_NAMES[provider]) or "").strip()
    key = str(config.get(PROVIDER_KEY_NAMES[provider]) or "").strip()
    if not key or not model:
        raise AIRouterError(f"{PROVIDER_LABELS[provider]} için API anahtarı veya model adı eksik.")
    if provider == "openai":
        output = _call_openai(config, model, instructions, input_data)
    elif provider == "anthropic":
        output = _call_anthropic(config, model, instructions, input_data)
    elif provider == "gemini":
        output = _call_gemini(config, model, instructions, input_data)
    else:
        if not str(config.get("AI_COMPAT_BASE_URL") or "").strip():
            raise AIRouterError("OpenAI uyumlu sağlayıcı için AI_COMPAT_BASE_URL eksik.")
        output = _call_openai(config, model, instructions, input_data, compatible=True)
    if not output.strip():
        raise AIRouterError(f"{PROVIDER_LABELS[provider]} boş yanıt döndürdü.")
    return AIResult(output_text=output, provider=provider, model=model)


def _candidate_order(config: dict[str, Any]) -> list[str]:
    available = configured_providers(config)
    if not available:
        raise AIRouterError("Hiçbir yapay zekâ sağlayıcısı hazır değil. Sağlayıcı API anahtarı ve model adını sunucu Secrets alanına tanımlayın.")
    selected = str(config.get("AI_PROVIDER") or "auto").strip().lower()
    if selected != "auto":
        if selected not in PROVIDER_LABELS:
            raise AIRouterError("AI_PROVIDER değeri openai, anthropic, gemini, openai_compatible veya auto olmalıdır.")
        if selected not in available:
            raise AIRouterError(f"Seçilen {PROVIDER_LABELS[selected]} sağlayıcısı yapılandırılmamış.")
        ordered = [selected] + [p for p in available if p != selected]
    else:
        ordered = available
    if not _truthy(config.get("AI_FALLBACK_ENABLED", "1")):
        return ordered[:1]
    return ordered


class AIRouterClient:
    """Drop-in subset for the app's existing client.responses.create call sites."""
    def __init__(self, config: dict[str, Any]):
        self.config = dict(config)
        self.responses = self
        self.provider_order = _candidate_order(self.config)
        self.model = str(self.config.get(PROVIDER_MODEL_NAMES[self.provider_order[0]]) or "")

    def create(self, model: str | None = None, instructions: str = "", input: Any = "", **kwargs: Any) -> AIResult:
        errors = []
        providers = self.provider_order
        for ix, provider in enumerate(providers):
            try:
                result = call_provider(self.config, provider, instructions, input)
                result.fallback_used = ix > 0
                return result
            except Exception as exc:
                errors.append(f"{PROVIDER_LABELS[provider]}: {str(exc)[:180]}")
        raise AIRouterError("Tüm yapılandırılmış AI sağlayıcıları başarısız oldu. " + " | ".join(errors))


def multi_provider_review(config: dict[str, Any], instructions: str, input_data: Any,
                          max_providers: int = 3) -> list[dict[str, str]]:
    """Call independently configured providers; partial results are retained."""
    try:
        limit = int(max_providers)
    except (TypeError, ValueError):
        limit = 3
    candidates = configured_providers(config)[:max(1, min(limit, 4))]
    results = []
    for provider in candidates:
        try:
            result = call_provider(config, provider, instructions, input_data)
            results.append({"provider": result.provider, "model": result.model,
                            "status": "success", "output_text": result.output_text, "error": ""})
        except Exception as exc:
            results.append({"provider": provider, "model": str(config.get(PROVIDER_MODEL_NAMES[provider]) or ""),
                            "status": "failed", "output_text": "", "error": str(exc)[:300]})
    if not candidates:
        results.append({"provider": "", "model": "", "status": "not_configured", "output_text": "",
                        "error": "Çoklu AI denetimi için en az bir sağlayıcı API anahtarı ve model adı tanımlanmalıdır."})
    return results
