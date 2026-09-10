from __future__ import annotations

import threading
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any

from core.capabilities import CapabilityDetector
from core.connectivity import ConnectivityDiagnostics


@dataclass(slots=True)
class WorkstationSession:
    host: str
    transport: str
    opened_at: str
    connectivity: dict[str, Any] = field(default_factory=dict)
    capabilities: dict[str, Any] = field(default_factory=dict)
    capability_error: str | None = None

    @property
    def ready(self) -> bool:
        return str(self.connectivity.get("state", "")).startswith("READY_")


class SessionManager:
    """Cache thread-safe de contexto lógico por estação."""

    def __init__(self, executor) -> None:
        self.executor = executor
        self.connectivity = ConnectivityDiagnostics(executor)
        self.capability_detector = CapabilityDetector(executor)
        self._sessions: dict[str, WorkstationSession] = {}
        self._guard = threading.RLock()

    def open(
        self,
        host: str,
        *,
        refresh: bool = False,
    ) -> WorkstationSession:
        key = host.casefold()
        with self._guard:
            if key in self._sessions and not refresh:
                return self._sessions[key]

        connectivity = self.connectivity.run(host)
        capabilities: dict[str, Any] = {}
        capability_error: str | None = None
        transport = str(connectivity.get("selected_transport") or "unknown")

        if str(connectivity.get("state", "")).startswith("READY_"):
            report = self.capability_detector.probe(host)
            capabilities = report.values
            capability_error = report.error
            if report.transport:
                transport = report.transport

        session = WorkstationSession(
            host=host,
            transport=transport,
            opened_at=datetime.now(timezone.utc).isoformat(),
            connectivity=connectivity,
            capabilities=capabilities,
            capability_error=capability_error,
        )
        with self._guard:
            self._sessions[key] = session
        return session

    def get(self, host: str) -> WorkstationSession | None:
        with self._guard:
            return self._sessions.get(host.casefold())

    def close(self, host: str) -> None:
        with self._guard:
            self._sessions.pop(host.casefold(), None)

    def clear(self) -> None:
        with self._guard:
            self._sessions.clear()
