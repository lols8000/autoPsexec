from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any

from core.jobs import OperationClass
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

    def parse(self, raw: Any) -> Any:
        if raw is None or (isinstance(raw, str) and not raw.strip()):
            if self.default is not None:
                return self.default
            if self.required:
                raise ValueError(f"Parâmetro obrigatório: {self.label}")
            return None

        if self.kind is ParameterKind.INTEGER:
            try:
                value = int(str(raw).strip())
            except ValueError as exc:
                raise ValueError(
                    f"{self.label} deve ser um número inteiro."
                ) from exc
            if self.min_value is not None and value < self.min_value:
                raise ValueError(
                    f"{self.label} deve ser >= {self.min_value}."
                )
            if self.max_value is not None and value > self.max_value:
                raise ValueError(
                    f"{self.label} deve ser <= {self.max_value}."
                )
            return value

        if self.kind is ParameterKind.BOOLEAN:
            if isinstance(raw, bool):
                return raw
            value = str(raw).strip().casefold()
            truthy = {"1", "true", "sim", "s", "yes", "y"}
            falsy = {"0", "false", "não", "nao", "n", "no"}
            if value in truthy:
                return True
            if value in falsy:
                return False
            raise ValueError(
                f"{self.label} deve ser SIM/NÃO."
            )

        value = str(raw).strip()
        if self.kind is ParameterKind.CHOICE:
            if value not in self.choices:
                allowed = ", ".join(self.choices)
                raise ValueError(
                    f"{self.label} deve ser um de: {allowed}."
                )
        return value


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
    requires_confirmation: bool = True
    requires_reboot: bool = False
    may_break_connectivity: bool = False
    destructive: bool = False
    parameters: tuple[ExecutionParameter, ...] = ()
    recommendation: str | None = None


@dataclass(slots=True)
class ExecutionRecord:
    action: ExecutionAction
    host: str
    parameters: dict[str, Any]
    remediation: RemediationResult

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
