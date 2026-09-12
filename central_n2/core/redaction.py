from __future__ import annotations

import re
from dataclasses import asdict, is_dataclass
from enum import Enum
from typing import Any

_KEY_RE = re.compile(
    r"(?i)(password|passwd|senha|secret|token|api[_-]?key|authorization|credential)"
)
_INLINE_RE = re.compile(
    r"(?i)\b(password|passwd|senha|secret|token|api[_-]?key|authorization)\b"
    r"\s*[:=]\s*([^\s,;]+)"
)
_BEARER_RE = re.compile(
    r"(?i)\bBearer\s+[A-Za-z0-9._~+/=-]+"
)


def redact_text(value: str) -> str:
    value = _INLINE_RE.sub(
        lambda match: f"{match.group(1)}=***",
        value,
    )
    return _BEARER_RE.sub("Bearer ***", value)


def redact(
    value: Any,
    *,
    key: str | None = None,
) -> Any:
    if key and _KEY_RE.search(key):
        return "***"

    if is_dataclass(value) and not isinstance(value, type):
        return redact(asdict(value), key=key)

    if isinstance(value, Enum):
        return redact(value.value, key=key)

    if isinstance(value, str):
        return redact_text(value)

    if isinstance(value, dict):
        return {
            item_key: redact(
                item_value,
                key=str(item_key),
            )
            for item_key, item_value in value.items()
        }

    if isinstance(value, list):
        return [redact(item) for item in value]

    if isinstance(value, tuple):
        return tuple(redact(item) for item in value)

    if isinstance(value, set):
        return [
            redact(item)
            for item in sorted(value, key=str)
        ]

    return value
