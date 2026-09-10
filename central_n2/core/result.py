from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class CommandResult:
    success: bool
    command: str
    host: str
    stdout: str = ""
    stderr: str = ""
    return_code: int = 0
    duration_ms: int = 0
    transport: str = "local"
    data: Any = None
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def indeterminate(self) -> bool:
        """True quando a ação pode ter chegado ao destino, mas não houve confirmação."""
        return bool(self.metadata.get("indeterminate"))

    def mark_indeterminate(self, reason: str) -> "CommandResult":
        self.metadata["indeterminate"] = True
        self.metadata["indeterminate_reason"] = reason
        return self

    @classmethod
    def failure(
        cls,
        host: str,
        command: str,
        message: str,
        *,
        return_code: int = 1,
        transport: str = "local",
    ) -> "CommandResult":
        return cls(
            success=False,
            command=command,
            host=host,
            stderr=message,
            return_code=return_code,
            transport=transport,
        )
