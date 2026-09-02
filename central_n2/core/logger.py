from __future__ import annotations

import json
import os
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from .result import CommandResult


class AuditLogger:
    def __init__(self, log_dir: str | Path = "logs") -> None:
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)

    def _path(self) -> Path:
        return self.log_dir / f"{datetime.now().strftime('%Y-%m-%d')}.jsonl"

    def log_result(self, action: str, result: CommandResult, **extra: Any) -> None:
        payload: Dict[str, Any] = {
            "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
            "operator": os.environ.get("USERNAME") or os.environ.get("USER") or "unknown",
            "action": action,
            **asdict(result),
        }
        payload.update(extra)
        with self._path().open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(payload, ensure_ascii=False, default=str) + "\n")

    def log_event(self, action: str, host: str, status: str, **extra: Any) -> None:
        payload = {
            "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
            "operator": os.environ.get("USERNAME") or os.environ.get("USER") or "unknown",
            "action": action,
            "host": host,
            "status": status,
            **extra,
        }
        with self._path().open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(payload, ensure_ascii=False, default=str) + "\n")
