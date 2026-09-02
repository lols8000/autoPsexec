from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Optional


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
    data: Optional[Any] = None
    metadata: Dict[str, Any] = field(default_factory=dict)

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
