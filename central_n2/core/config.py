from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any


def deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = deepcopy(value)
    return result


class ConfigLoader:
    """Carrega configuração pública e aplica override local não versionado."""

    def __init__(self, settings_path: str | Path) -> None:
        self.settings_path = Path(settings_path)
        self.local_path = self.settings_path.with_name("settings.local.json")
        self._settings: dict[str, Any] = {}
        self.reload()

    def reload(self) -> dict[str, Any]:
        base = json.loads(self.settings_path.read_text(encoding="utf-8"))
        override: dict[str, Any] = {}
        if self.local_path.exists():
            override = json.loads(self.local_path.read_text(encoding="utf-8"))
        self._settings = deep_merge(base, override)
        return deepcopy(self._settings)

    @property
    def settings(self) -> dict[str, Any]:
        return deepcopy(self._settings)

    def get(self, key: str, default: Any = None) -> Any:
        return self._settings.get(key, default)
