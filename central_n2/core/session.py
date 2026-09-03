from __future__ import annotations

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


class SessionManager:
    def __init__(self, executor) -> None:
        self.executor = executor
        self.connectivity = ConnectivityDiagnostics(executor)
        self.capability_detector = CapabilityDetector(executor)
        self._sessions: dict[str, WorkstationSession] = {}

    def open(self, host: str, *, refresh: bool = False) -> WorkstationSession:
        key = host.lower()
        if key in self._sessions and not refresh:
            return self._sessions[key]
        connectivity = self.connectivity.run(host)
        capabilities = self.capability_detector.probe(host)
        session = WorkstationSession(
            host=host,
            transport=connectivity.get("selected_transport", capabilities.transport),
            opened_at=datetime.now(timezone.utc).isoformat(),
            connectivity=connectivity,
            capabilities=capabilities.values,
        )
        self._sessions[key] = session
        return session

    def get(self, host: str) -> WorkstationSession | None:
        return self._sessions.get(host.lower())

    def close(self, host: str) -> None:
        self._sessions.pop(host.lower(), None)

    def clear(self) -> None:
        self._sessions.clear()
