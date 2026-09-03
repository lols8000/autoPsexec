from __future__ import annotations
import json
import os
import threading
import uuid
from contextlib import contextmanager
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterator
from .redaction import redact
from .result import CommandResult

class AuditLogger:
    def __init__(self, log_dir: str | Path = "logs") -> None:
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self._local = threading.local()

    def _path(self) -> Path:
        return self.log_dir / f"{datetime.now().strftime('%Y-%m-%d')}.jsonl"

    @contextmanager
    def bind(self, **context: Any) -> Iterator[None]:
        previous = getattr(self._local, "context", {})
        self._local.context = {**previous, **context}
        try:
            yield
        finally:
            self._local.context = previous

    def _base(self, action: str, host: str) -> Dict[str, Any]:
        context = dict(getattr(self._local, "context", {}))
        return {
            "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
            "operator": os.environ.get("USERNAME") or os.environ.get("USER") or "unknown",
            "correlation_id": context.pop("correlation_id", uuid.uuid4().hex[:16]),
            "action": action, "host": host, **context,
        }

    def _write(self, payload: Dict[str, Any]) -> None:
        with self._path().open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(redact(payload), ensure_ascii=False, default=str) + "\n")

    def log_result(self, action: str, result: CommandResult, **extra: Any) -> None:
        payload = self._base(action, result.host)
        payload.update(asdict(result)); payload.update(extra); self._write(payload)

    def log_event(self, action: str, host: str, status: str, **extra: Any) -> None:
        payload = self._base(action, host)
        payload.update({"status": status, **extra}); self._write(payload)
