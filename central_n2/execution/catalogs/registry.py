from __future__ import annotations

from core.jobs import OperationClass
from ..models import ExecutionAction, ExecutionParameter, ParameterKind, RiskLevel
from ..validators import command_field_true
from .common import ExecutionDependencies, _register


def register(registry, deps: ExecutionDependencies) -> None:
    registry_keys = deps.registry_actions.keys()
    if registry_keys:
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
            ),
            lambda host, p: deps.registry_actions.apply(
                host,
                p["registry_key"],
            ),
            validator=command_field_true(
                "Applied",
                pass_message="Ação de Registro aplicada.",
                fail_message="Ação de Registro não foi confirmada.",
            ),
        )
