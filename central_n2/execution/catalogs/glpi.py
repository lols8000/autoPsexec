from __future__ import annotations

from core.jobs import OperationClass
from ..models import ExecutionAction, RiskLevel
from .common import ExecutionDependencies, _glpi_running, _register


def register(registry, deps: ExecutionDependencies) -> None:
    for key, title, fn, risk in (
        (
            "glpi.restart",
            "Reiniciar GLPI Agent",
            deps.glpi.restart_service,
            RiskLevel.LOW,
        ),
        (
            "glpi.install_repair",
            "Instalar / reparar GLPI Agent",
            deps.glpi.install_or_repair,
            RiskLevel.MEDIUM,
        ),
        (
            "glpi.force_inventory",
            "Forçar inventário GLPI",
            deps.glpi.force_inventory,
            RiskLevel.LOW,
        ),
    ):
        _register(
            registry,
            ExecutionAction(
                key,
                title,
                "glpi",
                "GLPI Agent",
                title,
                OperationClass.HEAVY_WRITE,
                risk,
                "Altera ou aciona o agente de inventário.",
                900,
            ),
            lambda host, p, action=fn: action(host),
            before_probe=lambda host, p: deps.glpi.status(host),
            after_probe=lambda host, p: deps.glpi.status(host),
            validator=_glpi_running,
        )
