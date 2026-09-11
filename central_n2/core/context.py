from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass(slots=True)
class AttendanceContext:
    """Estado isolado de um único atendimento/estação."""

    correlation_id: str = field(default_factory=lambda: uuid.uuid4().hex[:16])
    host: str | None = None
    session: Any = None
    health_snapshot: dict[str, Any] | None = None
    diagnoses: list[Any] = field(default_factory=list)
    playbook: Any = None
    remediation: Any = None
    execution: Any = None
    report_path: Path | None = None

    @classmethod
    def start(cls, host: str | None = None, session: Any = None) -> "AttendanceContext":
        return cls(host=host, session=session)

    def bind(self, host: str, session: Any) -> None:
        """Vincula o contexto a um host novo e zera qualquer estado anterior."""
        self.correlation_id = uuid.uuid4().hex[:16]
        self.host = host
        self.session = session
        self.health_snapshot = None
        self.diagnoses.clear()
        self.playbook = None
        self.remediation = None
        self.execution = None
        self.report_path = None

    def belongs_to(self, host: str | None) -> bool:
        return bool(host and self.host and host.casefold() == self.host.casefold())
