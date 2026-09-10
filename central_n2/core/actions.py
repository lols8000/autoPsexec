from __future__ import annotations

from dataclasses import dataclass

from core.jobs import OperationClass


@dataclass(frozen=True, slots=True)
class ActionSpec:
    """Metadados operacionais independentes do texto apresentado na UI."""

    key: str
    title: str
    operation_class: OperationClass = OperationClass.READ_ONLY
    timeout_seconds: int | None = None
    requires_confirmation: bool = False

    @classmethod
    def read(
        cls,
        key: str,
        title: str,
        *,
        timeout_seconds: int | None = None,
    ) -> "ActionSpec":
        return cls(
            key=key,
            title=title,
            operation_class=OperationClass.READ_ONLY,
            timeout_seconds=timeout_seconds,
        )
