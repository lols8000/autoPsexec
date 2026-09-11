from __future__ import annotations

from core.jobs import OperationClass
from ..models import ExecutionAction, ExecutionParameter, ParameterKind, RiskLevel
from ..validators import service_running, service_stopped
from .common import ExecutionDependencies, _register, _startup_type


def register(registry, deps: ExecutionDependencies) -> None:
    service_name = (
        ExecutionParameter("service_name", "Nome do serviço", ParameterKind.TEXT),
    )
    for key, title, method, validator in (
        ("service.start", "Iniciar serviço", "start", service_running),
        ("service.stop", "Parar serviço", "stop", service_stopped),
        ("service.restart", "Reiniciar serviço", "restart", service_running),
    ):
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
                ExecutionParameter("service_name", "Nome do serviço", ParameterKind.TEXT),
                ExecutionParameter(
                    "startup_type",
                    "Tipo de inicialização",
                    ParameterKind.CHOICE,
                    choices=("Automatic", "Manual", "Disabled"),
                ),
            ),
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
    )
