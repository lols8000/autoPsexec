from __future__ import annotations

from core.jobs import OperationClass
from ..models import ExecutionAction, RiskLevel
from ..validators import command_completed
from .common import ExecutionDependencies, _register


def register(registry, deps: ExecutionDependencies) -> None:
    for key, title, fn, timeout, risk in (
        (
            "windows.sfc",
            "Executar SFC /scannow",
            deps.repair.sfc_scan,
            2400,
            RiskLevel.MEDIUM,
        ),
        (
            "windows.dism_restore",
            "Executar DISM RestoreHealth",
            deps.repair.dism_restorehealth,
            4200,
            RiskLevel.MEDIUM,
        ),
        (
            "windows.component_cleanup",
            "Limpar Component Store",
            deps.repair.component_cleanup,
            2400,
            RiskLevel.MEDIUM,
        ),
        (
            "windows.store_reset",
            "Resetar Microsoft Store",
            deps.repair.reset_store,
            600,
            RiskLevel.LOW,
        ),
    ):
        _register(
            registry,
            ExecutionAction(
                key,
                title,
                "windows",
                "Windows / Sistema",
                title,
                OperationClass.HEAVY_WRITE,
                risk,
                "Pode consumir CPU/disco e levar vários minutos.",
                timeout,
            ),
            lambda host, p, action=fn: action(host),
            validator=command_completed,
        )
