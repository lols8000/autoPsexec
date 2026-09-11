from __future__ import annotations

from typing import Any

from core.jobs import OperationClass
from core.result import CommandResult
from remediation import ValidationResult, ValidationStatus

from ..models import (
    ExecutionAction,
    ExecutionParameter,
    ParameterKind,
    RiskLevel,
)
from .common import ExecutionDependencies, _register


def _payload(value: Any) -> dict[str, Any]:
    if isinstance(value, CommandResult):
        return value.data if isinstance(value.data, dict) else {}
    return value if isinstance(value, dict) else {}


def register(registry, deps: ExecutionDependencies) -> None:
    registry_keys = deps.registry_actions.keys()
    if not registry_keys:
        return

    def validate_apply(
        before: Any,
        command: CommandResult,
        after: Any,
        parameters: dict[str, Any],
    ) -> ValidationResult:
        data = _payload(after)
        item = deps.registry_actions.catalog[
            parameters["registry_key"]
        ]
        mode = str(item.get("mode") or "set").casefold()

        if command.indeterminate:
            return ValidationResult(
                ValidationStatus.UNKNOWN,
                "A alteração de Registro teve resultado indeterminado.",
                data,
            )

        if mode == "remove":
            ok = data.get("Exists") is False
        else:
            ok = data.get("Exists") is True
            if ok and "value" in item:
                ok = str(data.get("Value")) == str(item.get("value"))

        if ok:
            return ValidationResult(
                ValidationStatus.PASS,
                "Estado do Registro confirmado após a ação.",
                data,
            )
        return ValidationResult(
            ValidationStatus.FAIL if not command.success else ValidationStatus.UNKNOWN,
            "Estado final do Registro não foi confirmado.",
            data,
        )

    def validate_rollback(
        before: Any,
        command: CommandResult,
        after: Any,
        parameters: dict[str, Any],
    ) -> ValidationResult:
        original = _payload(before)
        current = _payload(after)

        if command.indeterminate:
            return ValidationResult(
                ValidationStatus.UNKNOWN,
                "Rollback do Registro teve resultado indeterminado.",
                current,
            )

        expected_exists = original.get("Exists")
        if expected_exists is False and current.get("Exists") is False:
            return ValidationResult(
                ValidationStatus.PASS,
                "Ausência original do valor foi restaurada.",
                current,
            )
        if expected_exists is True:
            same = (
                current.get("Exists") is True
                and str(current.get("Value"))
                == str(original.get("Value"))
            )
            if same:
                return ValidationResult(
                    ValidationStatus.PASS,
                    "Valor original do Registro foi restaurado.",
                    current,
                )

        return ValidationResult(
            ValidationStatus.FAIL if not command.success else ValidationStatus.UNKNOWN,
            "Estado original do Registro não foi confirmado.",
            current,
        )

    _register(
        registry,
        ExecutionAction(
            "registry.apply",
            "Aplicar ação de Registro homologada",
            "registry",
            "Registro",
            "Executa somente alteração HKLM definida na configuração local.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Altera configuração persistente do Windows/aplicação.",
            300,
            parameters=(
                ExecutionParameter(
                    "registry_key",
                    "Ação homologada",
                    ParameterKind.CHOICE,
                    choices=registry_keys,
                ),
            ),
            idempotent=True,
            rollback_strategy=(
                "Restaurar o valor/ausência observado antes da ação."
            ),
            tags=("registro", "registry", "hklm", "política"),
        ),
        lambda host, p: deps.registry_actions.apply(
            host,
            p["registry_key"],
        ),
        before_probe=lambda host, p: deps.registry_actions.inspect(
            host,
            p["registry_key"],
        ),
        after_probe=lambda host, p: deps.registry_actions.inspect(
            host,
            p["registry_key"],
        ),
        validator=validate_apply,
        rollback_handler=lambda host, p, before: (
            deps.registry_actions.rollback(
                host,
                p["registry_key"],
                before,
            )
        ),
        rollback_validator=validate_rollback,
    )
