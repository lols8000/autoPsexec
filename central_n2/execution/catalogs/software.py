from __future__ import annotations

from core.jobs import OperationClass

from ..models import (
    ExecutionAction,
    ExecutionParameter,
    ParameterKind,
    PrivilegeLevel,
    RiskLevel,
)
from ..validators import command_completed
from .common import ExecutionDependencies, _register


def register(registry, deps: ExecutionDependencies) -> None:
    software_keys = tuple(sorted(deps.software.catalog))
    if not software_keys:
        return

    software_parameter = ExecutionParameter(
        "software_key",
        "Software homologado",
        ParameterKind.CHOICE,
        choices=software_keys,
    )
    for key, title, handler, destructive in (
        (
            "software.install",
            "Instalar software do catálogo",
            deps.software.install_catalog_item,
            False,
        ),
        (
            "software.upgrade",
            "Atualizar software do catálogo",
            deps.software.upgrade_catalog_item,
            False,
        ),
        (
            "software.uninstall",
            "Remover software do catálogo",
            deps.software.uninstall_catalog_item,
            True,
        ),
    ):
        _register(
            registry,
            ExecutionAction(
                key,
                title,
                "software",
                "Software",
                "Executa Winget somente para itens homologados no catálogo.",
                OperationClass.HEAVY_WRITE,
                RiskLevel.HIGH if destructive else RiskLevel.MEDIUM,
                "Altera software instalado na estação.",
                900,
                destructive=destructive,
                parameters=(software_parameter,),
                allowed_transports=("local", "winrm"),
                required_capabilities=("Winget",),
                required_privilege=PrivilegeLevel.USER_CONTEXT,
                tags=("software", "winget", key.rsplit(".", 1)[-1]),
            ),
            lambda host, p, fn=handler: fn(
                host,
                p["software_key"],
            ),
            validator=command_completed,
        )
