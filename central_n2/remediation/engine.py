from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from enum import Enum
from typing import Any, Callable

from core.result import CommandResult


class ValidationStatus(str, Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    UNKNOWN = "UNKNOWN"


@dataclass(frozen=True, slots=True)
class RemediationSpec:
    key: str
    title: str
    impact: str
    requires_confirmation: bool = True
    requires_reboot: bool = False
    disruptive: bool = False
    rollback: str | None = None


@dataclass(slots=True)
class ValidationResult:
    status: ValidationStatus
    message: str
    evidence: Any = None

    @property
    def success(self) -> bool | None:
        if self.status is ValidationStatus.PASS:
            return True
        if self.status is ValidationStatus.FAIL:
            return False
        return None


@dataclass(slots=True)
class RemediationResult:
    spec: RemediationSpec
    host: str
    before: Any
    command_result: CommandResult
    after: Any
    validation: ValidationResult
    started_at: str
    finished_at: str


Validator = Callable[[Any, CommandResult, Any], ValidationResult]
Probe = Callable[[str], Any]


class RemediationEngine:
    def execute(
        self,
        host: str,
        spec: RemediationSpec,
        action: Callable[[str], CommandResult],
        *,
        snapshotter: Probe | None = None,
        before_probe: Probe | None = None,
        after_probe: Probe | None = None,
        validator: Validator | None = None,
    ) -> RemediationResult:
        started = datetime.now().astimezone().isoformat(timespec="seconds")
        before_fn = before_probe or snapshotter
        after_fn = after_probe or snapshotter

        before = before_fn(host) if before_fn else None
        result = action(host)

        after = None
        if after_fn:
            try:
                after = after_fn(host)
            except Exception as exc:
                after = {"probe_error": f"{type(exc).__name__}: {exc}"}

        if validator:
            try:
                validation = validator(before, result, after)
            except Exception as exc:
                validation = ValidationResult(
                    ValidationStatus.UNKNOWN,
                    f"Falha ao validar remediação: {type(exc).__name__}: {exc}",
                    after,
                )
        elif result.indeterminate:
            validation = ValidationResult(
                ValidationStatus.UNKNOWN,
                "A execução teve resultado indeterminado e não possui validador específico.",
                after,
            )
        elif result.success:
            validation = ValidationResult(
                ValidationStatus.PASS,
                "Comando concluído; não há validador específico configurado.",
                after,
            )
        else:
            validation = ValidationResult(
                ValidationStatus.FAIL,
                "O comando de remediação falhou.",
                after,
            )

        return RemediationResult(
            spec=spec,
            host=host,
            before=before,
            command_result=result,
            after=after,
            validation=validation,
            started_at=started,
            finished_at=datetime.now().astimezone().isoformat(
                timespec="seconds"
            ),
        )
