"""Server-only deployment settings. Never copied into the browser bundle."""
from __future__ import annotations

import json
import os
from pathlib import Path
from urllib.parse import urlparse

DEFAULT_BASE_URL = "https://api.deepseek.com/v1"
DEFAULT_MODEL = "请填写服务商的模型 ID"


def load_config(root: Path) -> None:
    path = Path(os.environ.get("ZHAI_CONFIG_FILE", str(root / "private-config.json")))
    if path.is_file():
        values = json.loads(path.read_text(encoding="utf-8-sig"))
        allowed = {"ZHAI_LLM_BASE_URL", "ZHAI_LLM_MODEL", "ZHAI_LLM_NAME", "ZHAI_LLM_API_KEY", "ZHAI_ACCESS_USER", "ZHAI_ACCESS_PASSWORD"}
        for name, value in values.items():
            if name not in allowed or not isinstance(value, str):
                raise ValueError("服务器配置包含不支持的字段")
            os.environ.setdefault(name, value)
    os.environ.setdefault("ZHAI_LLM_BASE_URL", DEFAULT_BASE_URL)
    os.environ.setdefault("ZHAI_LLM_MODEL", DEFAULT_MODEL)
    os.environ.setdefault("ZHAI_LLM_NAME", "自配模型 · DeepSeek")


def model_settings(payload: dict) -> tuple[str, str, str]:
    default_url = os.environ.get("ZHAI_LLM_BASE_URL", DEFAULT_BASE_URL).rstrip("/")
    url = str(payload.get("baseUrl") or default_url).rstrip("/")
    supplied_key = str(payload.get("apiKey") or "").strip()
    parsed = urlparse(url)
    if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise ValueError("模型服务地址必须是有效的 HTTPS 地址")
    # A changed endpoint must never receive the institution's server-side key.
    if not supplied_key and url != default_url:
        raise ValueError("使用自定义接口时，请同时填写你自己的 API Key；单位密钥只用于预设算力地址")
    key = supplied_key or os.environ.get("ZHAI_LLM_API_KEY", "")
    model = str(payload.get("model") or os.environ.get("ZHAI_LLM_MODEL", DEFAULT_MODEL))
    return url, key, model


def safe_error(exc: Exception) -> str:
    message = str(exc)
    for name in ("ZHAI_LLM_API_KEY", "ZHAI_ACCESS_PASSWORD"):
        value = os.environ.get(name)
        if value:
            message = message.replace(value, "[已隐藏]")
    return message
