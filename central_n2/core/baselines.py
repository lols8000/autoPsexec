from __future__ import annotations
import json
from pathlib import Path
from typing import Any
from core.config import deep_merge

class BaselineRepository:
    def __init__(self, directory: str | Path) -> None:
        self.directory = Path(directory)

    def available(self) -> list[str]:
        return sorted(p.stem.upper() for p in self.directory.glob("*.json"))

    def load(self, profile: str, *, fallback: str = "DEFAULT") -> dict[str, Any]:
        requested = self.directory / f"{profile.upper()}.json"
        default = self.directory / f"{fallback.upper()}.json"
        base = json.loads(default.read_text(encoding="utf-8")) if default.exists() else {}
        if requested.exists() and requested != default:
            return deep_merge(base, json.loads(requested.read_text(encoding="utf-8")))
        return base
