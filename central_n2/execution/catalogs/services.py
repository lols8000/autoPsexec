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
    SelectorKind,
)
from ..validators import service_running, service_stopped
from .common import ExecutionDependencies, _register, _startup_type


def _payload(value: Any) -> dict[str, Any]:
    if isinstance(value, CommandResult):
        return value.data if isinstance(value.data, dict) else {}
    return value if isinstance(value, dict) else {}


def _rollback_service(
    deps: ExecutionDependencies,
    host: str,
    parameters: dict[str, Any],
    before: Any,
) -> CommandResult:
    original = _payload(before)
    status = str(original.get("Status") or "").casefold()
    if status == "running":
        action = "start"
    elif status == "stopped":
        action = "stop"
    else:
        return CommandResult.failure(
            host,
            "service.rollback",
            "Estado original do serviço não está disponível.",
        )
    return deps.system.service_action(
        host,
        parameters["service_name"],
        action,
    )


def _validate_service_rollback(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    original = _payload(before)
    current = _payload(after)
    expected = str(original.get("Status") or "").casefold()
    actual = str(current.get("Status") or "").casefold()

    if command.indeterminate:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "Rollback teve resultado indeterminado.",
            current,
        )
    if expected and actual == expected:
        return ValidationResult(
            ValidationStatus.PASS,
            f"Estado original do serviço restaurado: {original.get('Status')}.",
            current,
        )
    return ValidationResult(
        ValidationStatus.FAIL if not command.success else ValidationStatus.UNKNOWN,
        "Estado original do serviço não foi confirmado.",
        current,
    )


def _rollback_startup(
    deps: ExecutionDependencies,
    host: str,
    parameters: dict[str, Any],
    before: Any,
) -> CommandResult:
    original = _payload(before)
    startup = str(original.get("StartType") or "")
    if startup not in {"Automatic", "Manual", "Disabled"}:
        return CommandResult.failure(
            host,
            "service.startup.rollback",
            "StartType original não está disponível.",
        )
    return deps.system.set_service_startup(
        host,
        parameters["service_name"],
        startup,
    )


def _validate_startup_rollback(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    original = _payload(before)
    current = _payload(after)
    expected = str(original.get("StartType") or "")
    actual = str(current.get("StartType") or "")

    if command.indeterminate:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "Rollback de StartType teve resultado indeterminado.",
            current,
        )
    if expected and actual.casefold() == expected.casefold():
        return ValidationResult(
            ValidationStatus.PASS,
            f"StartType original restaurado: {expected}.",
            current,
        )
    return ValidationResult(
        ValidationStatus.FAIL if not command.success else ValidationStatus.UNKNOWN,
        "StartType original não foi confirmado.",
        current,
    )


def _service_rollback_handler(deps: ExecutionDependencies):
    def handler(
        host: str,
        parameters: dict[str, Any],
        before: Any,
    ) -> CommandResult:
        return _rollback_service(
            deps,
            host,
            parameters,
            before,
        )

    return handler


def register(registry, deps: ExecutionDependencies) -> None:
    service_name = (
        ExecutionParameter(
            "service_name",
            "Serviço",
            ParameterKind.TEXT,
            selector=SelectorKind.SERVICE,
            help_text="Selecione um serviço inventariado ou informe o nome técnico.",
        ),
    )

    for key, title, method, validator in (
        ("service.start", "Iniciar serviço", "start", service_running),
        ("service.stop", "Parar serviço", "stop", service_stopped),
        ("service.restart", "Reiniciar serviço", "restart", service_running),
    ):
        rollback_handler = None
        rollback_validator = None
        rollback_strategy = None
        if method in {"start", "stop"}:
            rollback_handler = _service_rollback_handler(deps)
            rollback_validator = _validate_service_rollback
            rollback_strategy = "Restaurar o estado Running/Stopped observado antes da ação."

        _register(
            registry,
            ExecutionAction(
                key,
                title,
                "services",
                "Serviços",
                f"{title} e valida o estado final.",
                OperationClass.LIGHT_WRITE,
                RiskLevel.MEDIUM,
                "Pode afetar aplicações dependentes do serviço.",
                180,
                destructive=method == "stop",
                parameters=service_name,
                idempotent=method in {"start", "stop"},
                rollback_strategy=rollback_strategy,
                tags=("serviço", "service", method),
            ),
            lambda host, p, action=method: deps.system.service_action(
                host,
                p["service_name"],
                action,
            ),
            before_probe=lambda host, p: deps.system.service_status(
                host,
                p["service_name"],
            ),
            after_probe=lambda host, p: deps.system.service_status(
                host,
                p["service_name"],
            ),
            validator=validator,
            rollback_handler=rollback_handler,
            rollback_validator=rollback_validator,
        )

    _register(
        registry,
        ExecutionAction(
            "service.startup",
            "Alterar tipo de inicialização",
            "services",
            "Serviços",
            "Altera Automatic/Manual/Disabled para um serviço.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.HIGH,
            "Uma configuração incorreta pode impedir serviço essencial no próximo boot.",
            180,
            parameters=(
                ExecutionParameter(
                    "service_name",
                    "Serviço",
                    ParameterKind.TEXT,
                    selector=SelectorKind.SERVICE,
                ),
                ExecutionParameter(
                    "startup_type",
                    "Tipo de inicialização",
                    ParameterKind.CHOICE,
                    choices=("Automatic", "Manual", "Disabled"),
                ),
            ),
            idempotent=True,
            rollback_strategy="Restaurar o StartType observado antes da ação.",
            tags=("serviço", "startup", "inicialização"),
        ),
        lambda host, p: deps.system.set_service_startup(
            host,
            p["service_name"],
            p["startup_type"],
        ),
        before_probe=lambda host, p: deps.system.service_status(
            host,
            p["service_name"],
        ),
        after_probe=lambda host, p: deps.system.service_status(
            host,
            p["service_name"],
        ),
        validator=_startup_type,
        rollback_handler=lambda host, p, before: _rollback_startup(
            deps,
            host,
            p,
            before,
        ),
        rollback_validator=_validate_startup_rollback,
    )
