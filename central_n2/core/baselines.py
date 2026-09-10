from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from core.config import deep_merge


class BaselineRepository:
    def __init__(self, directory: str | Path) -> None:
        self.directory = Path(directory)

    def available(self) -> list[str]:
        return sorted(
            path.stem.upper()
            for path in self.directory.glob("*.json")
        )

    def load(
        self,
        profile: str,
        *,
        fallback: str = "DEFAULT",
    ) -> dict[str, Any]:
        requested = self.directory / f"{profile.upper()}.json"
        default = self.directory / f"{fallback.upper()}.json"
        base = (
            json.loads(default.read_text(encoding="utf-8"))
            if default.exists()
            else {}
        )
        if requested.exists() and requested != default:
            return deep_merge(
                base,
                json.loads(requested.read_text(encoding="utf-8")),
            )
        return base

    def resolve(
        self,
        profile: str,
        compliance_config: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Resolve DEFAULT, perfil e overrides explícitos nessa ordem.

        Chaves legadas fora de overrides só são aplicadas quando o perfil
        ativo é o perfil configurado. Isso preserva compatibilidade sem anular
        perfis escolhidos interativamente.
        """
        config = compliance_config or {}
        selected = profile.upper()
        effective = self.load(selected)
        overrides = dict(config.get("overrides", {}))

        configured = str(config.get("profile", "DEFAULT")).upper()
        if selected == configured:
            legacy = {
                key: value
                for key, value in config.items()
                if key not in {"profile", "overrides"}
            }
            overrides = deep_merge(legacy, overrides)

        return deep_merge(effective, overrides)
