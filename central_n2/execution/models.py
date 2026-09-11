from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any

from core.jobs import OperationClass
from core.redaction import redact
from core.result import CommandResult
from remediation import RemediationResult


class RiskLevel(str, Enum):
    LOW = "LOW"
    MEDIUM = "MEDIUM"
    HIGH = "HIGH"
    CRITICAL = "CRITICAL"


class ParameterKind(str, Enum):
    TEXT = "TEXT"
    INTEGER = "INTEGER"
    CHOICE = "CHOICE"
    BOOLEAN = "BOOLEAN"


class SelectorKind(str, Enum):
    PROCESS = "PROCESS"
    SERVICE = "SERVICE"
    ADAPTER = "ADAPTER"
    PRINTER = "PRINTER"
    PROFILE = "PROFILE"
    DEVICE = "DEVICE"
    SESSION = "SESSION"


class PrivilegeLevel(str, Enum):
    ANY = "ANY"
    ADMIN = "ADMIN"
    SYSTEM = "SYSTEM"
    USER_CONTEXT = "USER_CONTEXT"


class DisconnectMode(str, Enum):
    NONE = "NONE"
    TEMPORARY = "TEMPORARY"
    TERMINAL = "TERMINAL"


class RetryPolicy(str, Enum):
    NEVER = "NEVER"
    PRE_EXECUTION_ONLY = "PRE_EXECUTION_ONLY"
    SAFE_TRANSIENT = "SAFE_TRANSIENT"


@dataclass(frozen=True, slots=True)
class ExecutionParameter:
    key: str
    label: str
    kind: ParameterKind = ParameterKind.TEXT
    required: bool = True
    default: Any = None
    choices: tuple[str, ...] = ()
    min_value: int | None = None
    max_value: int | None = None
    sensitive: bool = False
    help_text: str | None = None
    selector: SelectorKind | None = None

    def parse(self, raw: Any) -> Any:
        if raw is None or (isinstance(raw, str) and not raw.strip()):
            if self.default is not None:
                return self.default
            if self.required:
                raise ValueError(f"Parâmetro obrigatório: {self.label}")
            return None

        if self.kind is ParameterKind.INTEGER:
            try:
                int_value = int(str(raw).strip())
            except ValueError as exc:
                raise ValueError(
                    f"{self.label} deve ser um número inteiro."
                ) from exc
            if self.min_value is not None and int_value < self.min_value:
                raise ValueError(
                    f"{self.label} deve ser >= {self.min_value}."
                )
            if self.max_value is not None and int_value > self.max_value:
                raise ValueError(
                    f"{self.label} deve ser <= {self.max_value}."
                )
            return int_value

        if self.kind is ParameterKind.BOOLEAN:
            if isinstance(raw, bool):
                return raw
            boolean_text = str(raw).strip().casefold()
            truthy = {"1", "true", "sim", "s", "yes", "y"}
            falsy = {"0", "false", "não", "nao", "n", "no"}
            if boolean_text in truthy:
                return True
            if boolean_text in falsy:
                return False
            raise ValueError(
                f"{self.label} deve ser SIM/NÃO."
            )

        text_value = str(raw).strip()
        if self.kind is ParameterKind.CHOICE:
            if text_value not in self.choices:
                allowed = ", ".join(self.choices)
                raise ValueError(
                    f"{self.label} deve ser um de: {allowed}."
                )
        return text_value


@dataclass(frozen=True, slots=True)
class ExecutionAction:
    key: str
    title: str
    category: str
    category_label: str
    description: str
    operation_class: OperationClass
    risk: RiskLevel
    impact: str
    timeout_seconds: int = 600
    requires_confirmation: bool = True
    requires_reboot: bool = False
    may_break_connectivity: bool = False
    destructive: bool = False
    parameters: tuple[ExecutionParameter, ...] = ()
    recommendation: str | None = None
    action_version: int = 1
    idempotent: bool = False
    retry_policy: RetryPolicy = RetryPolicy.PRE_EXECUTION_ONLY
    allowed_transports: tuple[str, ...] = (
        "local",
        "winrm",
        "psexec",
    )
    required_capabilities: tuple[str, ...] = ()
    required_privilege: PrivilegeLevel = PrivilegeLevel.ADMIN
    disconnect_mode: DisconnectMode = DisconnectMode.NONE
    recovery_timeout_seconds: int = 180
    recovery_delay_seconds: int = 3
    rollback_strategy: str | None = None
    tags: tuple[str, ...] = ()


@dataclass(slots=True)
class RecoveryResult:
    attempted: bool
    ready: bool
    attempts: int = 0
    elapsed_seconds: float = 0.0
    transport: str | None = None
    state: str | None = None
    error: str | None = None


@dataclass(slots=True)
class ExecutionPlan:
    action: ExecutionAction
    host: str
    parameters: dict[str, Any]
    policy: Any

    @property
    def allowed(self) -> bool:
        return bool(getattr(self.policy, "allowed", False))


@dataclass(slots=True)
class ExecutionRecord:
    action: ExecutionAction
    host: str
    parameters: dict[str, Any]
    remediation: RemediationResult
    policy: Any = None
    started_at: str | None = None
    finished_at: str | None = None
    duration_ms: int = 0
    operator: str | None = None
    recovery: Any = None
    rollback_available: bool = False
    rollback_result: CommandResult | None = None
    execution_id: int | None = None

    @property
    def public_parameters(self) -> dict[str, Any]:
        sensitive = {
            item.key
            for item in self.action.parameters
            if item.sensitive
        }
        return {
            key: "***" if key in sensitive else value
            for key, value in self.parameters.items()
        }


    def audit_payload(self) -> dict[str, Any]:
        command = self.remediation.command_result
        checks = []
        for item in getattr(self.policy, "checks", []):
            checks.append(
                {
                    "key": item.key,
                    "state": item.state.value,
                    "message": item.message,
                    "evidence": redact(item.evidence),
                }
            )

        recovery = None
        if self.recovery is not None:
            recovery = {
                "attempted": self.recovery.attempted,
                "ready": self.recovery.ready,
                "attempts": self.recovery.attempts,
                "elapsed_seconds": self.recovery.elapsed_seconds,
                "transport": self.recovery.transport,
                "state": self.recovery.state,
                "error": redact(self.recovery.error),
            }

        return {
            "action": {
                "key": self.action.key,
                "version": self.action.action_version,
                "title": self.action.title,
                "category": self.action.category,
                "risk": self.action.risk.value,
                "operation_class": self.action.operation_class.value,
                "disconnect_mode": self.action.disconnect_mode.value,
                "retry_policy": self.action.retry_policy.value,
                "idempotent": self.action.idempotent,
                "rollback_strategy": self.action.rollback_strategy,
            },
            "parameters": redact(self.public_parameters),
            "policy": {
                "allowed": bool(getattr(self.policy, "allowed", False)),
                "checks": checks,
            },
            "command": {
                "success": command.success,
                "transport": command.transport,
                "return_code": command.return_code,
                "duration_ms": command.duration_ms,
                "indeterminate": command.indeterminate,
                "error": redact(command.stderr),
                "data": redact(command.data),
            },
            "before": redact(self.remediation.before),
            "after": redact(self.remediation.after),
            "validation": {
                "status": self.remediation.validation.status.value,
                "message": self.remediation.validation.message,
                "evidence": redact(self.remediation.validation.evidence),
            },
            "started_at": self.started_at,
            "finished_at": self.finished_at,
            "duration_ms": self.duration_ms,
            "operator": self.operator,
            "recovery": recovery,
            "rollback_available": self.rollback_available,
        }


@dataclass(slots=True)
class ExecutionRollbackRecord:
    action_key: str
    host: str
    result: CommandResult
    validation: Any = None
    started_at: str | None = None
    finished_at: str | None = None
    duration_ms: int = 0
