from __future__ import annotations

from core.jobs import OperationClass
from remediation import validate_cleanup

from ..models import ExecutionAction, RiskLevel
from .common import ExecutionDependencies, _register, _wrap_three_arg


def register(registry, deps: ExecutionDependencies) -> None:
    _register(
        registry,
        ExecutionAction(
            "disk.cleanup_safe",
            "Limpeza segura de temporários",
            "disk",
            "Disco / Limpeza",
            "Limpa TEMP do usuário de execução e Windows Temp; não toca Downloads/Lixeira.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.MEDIUM,
            "Remove somente áreas temporárias previstas.",
            600,
        ),
        lambda host, p: deps.disk.cleanup_safe(host),
        before_probe=lambda host, p: deps.disk.cleanup_estimate(host),
        after_probe=lambda host, p: deps.disk.cleanup_estimate(host),
        validator=_wrap_three_arg(validate_cleanup),
    )
