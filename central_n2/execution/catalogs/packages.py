from __future__ import annotations

from core.jobs import OperationClass
from ..models import ExecutionAction, ExecutionParameter, ParameterKind, RiskLevel
from ..validators import command_completed
from .common import ExecutionDependencies, _register


def register(registry, deps: ExecutionDependencies) -> None:
    package_keys = deps.packages.keys()
    if package_keys:
        _register(
            registry,
            ExecutionAction(
                "package.install",
                "Instalar pacote corporativo homologado",
                "packages",
                "Pacotes corporativos",
                "Copia e executa somente pacote definido na configuração local.",
                OperationClass.HEAVY_WRITE,
                RiskLevel.HIGH,
                "Instala software/script corporativo com argumentos homologados.",
                2400,
                parameters=(
                    ExecutionParameter(
                        "package_key",
                        "Pacote",
                        ParameterKind.CHOICE,
                        choices=package_keys,
                    ),
                ),
            ),
            lambda host, p: deps.packages.install(
                host,
                p["package_key"],
            ),
            validator=command_completed,
        )
