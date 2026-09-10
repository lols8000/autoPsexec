from __future__ import annotations

import json
import os
import threading
import uuid
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator

from .redaction import redact
from .result import CommandResult


class AuditLogger:
    """Auditoria compacta e thread-safe.

    Payloads completos ficam desabilitados por padrão para evitar logs enormes e
    exposição desnecessária de comandos/saídas. O modo verbose é opt-in local.
    """

    def __init__(
        self,
        log_dir: str | Path = "logs",
        *,
        verbose_payloads: bool = False,
        max_error_chars: int = 4000,
    ) -> None:
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.verbose_payloads = bool(verbose_payloads)
        self.max_error_chars = max(256, int(max_error_chars))
        self._local = threading.local()
        self._write_lock = threading.Lock()

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

    def _base(self, action: str, host: str) -> dict[str, Any]:
        context = dict(getattr(self._local, "context", {}))
        return {
            "timestamp": datetime.now().astimezone().isoformat(
                timespec="seconds"
            ),
            "operator": (
                os.environ.get("USERNAME")
                or os.environ.get("USER")
                or "unknown"
            ),
            "correlation_id": context.pop(
                "correlation_id",
                uuid.uuid4().hex[:16],
            ),
            "action": action,
            "host": host,
            **context,
        }

    def _write(self, payload: dict[str, Any]) -> None:
        encoded = json.dumps(
            redact(payload),
            ensure_ascii=False,
            default=str,
        )
        with self._write_lock:
            with self._path().open("a", encoding="utf-8") as handle:
                handle.write(encoded + "\n")

    def log_result(
        self,
        action: str,
        result: CommandResult,
        **extra: Any,
    ) -> None:
        payload = self._base(action, result.host)
        payload.update(
            {
                "success": result.success,
                "return_code": result.return_code,
                "duration_ms": result.duration_ms,
                "transport": result.transport,
                "indeterminate": result.indeterminate,
                "metadata": result.metadata,
            }
        )

        if result.stderr:
            payload["error"] = result.stderr[: self.max_error_chars]

        if self.verbose_payloads:
            payload["command"] = result.command
            payload["stdout"] = result.stdout
            payload["stderr"] = result.stderr
            payload["data"] = result.data

        payload.update(extra)
        self._write(payload)

    def log_event(
        self,
        action: str,
        host: str,
        status: str,
        **extra: Any,
    ) -> None:
        payload = self._base(action, host)
        payload.update({"status": status, **extra})
        self._write(payload)
