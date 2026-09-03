from __future__ import annotations
import re
from typing import Any

_KEY_RE = re.compile(r"(?i)(password|passwd|senha|secret|token|api[_-]?key|authorization|credential)")
_INLINE_RE = re.compile(r"(?i)\b(password|passwd|senha|secret|token|api[_-]?key|authorization)\b\s*[:=]\s*([^\s,;]+)")
_BEARER_RE = re.compile(r"(?i)\bBearer\s+[A-Za-z0-9._~+/=-]+")

def redact_text(value: str) -> str:
    value = _INLINE_RE.sub(lambda m: f"{m.group(1)}=***", value)
    return _BEARER_RE.sub("Bearer ***", value)

def redact(value: Any, *, key: str | None = None) -> Any:
    if key and _KEY_RE.search(key):
        return "***"
    if isinstance(value, str):
        return redact_text(value)
    if isinstance(value, dict):
        return {k: redact(v, key=str(k)) for k, v in value.items()}
    if isinstance(value, list):
        return [redact(v) for v in value]
    if isinstance(value, tuple):
        return tuple(redact(v) for v in value)
    return value
