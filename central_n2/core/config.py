from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path
from typing import Any


class ConfigError(RuntimeError):
    pass


def deep_merge(
    base: dict[str, Any],
    override: dict[str, Any],
) -> dict[str, Any]:
    result = deepcopy(base)
    for key, value in override.items():
        if (
            isinstance(value, dict)
            and isinstance(result.get(key), dict)
        ):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = deepcopy(value)
    return result


def _positive_int(
    mapping: dict[str, Any],
    key: str,
    default: int,
) -> int:
    try:
        value = int(mapping.get(key, default))
    except (TypeError, ValueError) as exc:
        raise ConfigError(
            f"Configuração '{key}' deve ser inteira."
        ) from exc
    if value <= 0:
        raise ConfigError(
            f"Configuração '{key}' deve ser maior que zero."
        )
    return value


def validate_settings(
    settings: dict[str, Any],
) -> dict[str, Any]:
    if not isinstance(settings, dict):
        raise ConfigError(
            "O arquivo de configuração deve conter um objeto JSON."
        )

    _positive_int(
        settings,
        "timeout_seconds",
        60,
    )

    runtime = settings.get("runtime", {})
    if not isinstance(runtime, dict):
        raise ConfigError("'runtime' deve ser um objeto JSON.")
    _positive_int(runtime, "max_workers", 6)
    _positive_int(runtime, "max_batch_workers", 5)
    _positive_int(runtime, "retry_attempts", 2)

    ui = settings.get("ui", {})
    if not isinstance(ui, dict):
        raise ConfigError("'ui' deve ser um objeto JSON.")
    _positive_int(
        ui,
        "long_operation_timeout_seconds",
        3600,
    )

    compliance = settings.get("compliance", {})
    if not isinstance(compliance, dict):
        raise ConfigError("'compliance' deve ser um objeto JSON.")
    overrides = compliance.get("overrides", {})
    if not isinstance(overrides, dict):
        raise ConfigError(
            "'compliance.overrides' deve ser um objeto JSON."
        )

    return deepcopy(settings)


class ConfigLoader:
    """Carrega settings.json e aplica settings.local.json uma única vez."""

    def __init__(self, settings_path: str | Path) -> None:
        self.settings_path = Path(settings_path)
        self.local_path = self.settings_path.with_name(
            "settings.local.json"
        )
        self._settings: dict[str, Any] = {}
        self.reload()

    @staticmethod
    def _read_json(path: Path) -> dict[str, Any]:
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except FileNotFoundError as exc:
            raise ConfigError(
                f"Arquivo de configuração não encontrado: {path}"
            ) from exc
        except json.JSONDecodeError as exc:
            raise ConfigError(
                f"JSON inválido em {path.name}, linha "
                f"{exc.lineno}, coluna {exc.colno}: {exc.msg}"
            ) from exc
        if not isinstance(raw, dict):
            raise ConfigError(
                f"{path.name} deve conter um objeto JSON."
            )
        return raw

    def reload(self) -> dict[str, Any]:
        base = self._read_json(self.settings_path)
        override: dict[str, Any] = {}
        if self.local_path.exists():
            override = self._read_json(self.local_path)
        self._settings = validate_settings(
            deep_merge(base, override)
        )
        return deepcopy(self._settings)

    @property
    def settings(self) -> dict[str, Any]:
        return deepcopy(self._settings)

    def get(self, key: str, default: Any = None) -> Any:
        return deepcopy(self._settings.get(key, default))
