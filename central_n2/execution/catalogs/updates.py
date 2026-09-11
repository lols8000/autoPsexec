from __future__ import annotations

from core.jobs import OperationClass
from remediation import validate_windows_update_reset

from ..models import ExecutionAction, RiskLevel
from ..validators import command_field_true
from .common import ExecutionDependencies, _register, _update_install, _wrap_three_arg


def register(registry, deps: ExecutionDependencies) -> None:
    _register(
        registry,
        ExecutionAction(
            "updates.scan",
            "Disparar busca por atualizações",
            "updates",
            "Windows Update",
            "Solicita nova busca pelo UsoClient.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Baixo impacto.",
            180,
        ),
        lambda host, p: deps.updates.trigger_scan(host),
        validator=command_field_true(
            "ScanTriggered",
            pass_message="Busca por atualizações foi disparada.",
            fail_message="Busca não foi confirmada.",
        ),
    )
    _register(
        registry,
        ExecutionAction(
            "updates.install_pending",
            "Baixar e instalar atualizações pendentes",
            "updates",
            "Windows Update",
            "Usa Microsoft.Update.Session; não reinicia automaticamente.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Pode alterar componentes do Windows e exigir reinicialização.",
            7800,
            requires_reboot=True,
        ),
        lambda host, p: deps.updates.install_pending(host),
        before_probe=lambda host, p: deps.updates.status(host),
        after_probe=lambda host, p: deps.updates.status(host),
        validator=_update_install,
    )
    _register(
        registry,
        ExecutionAction(
            "updates.reset",
            "Resetar componentes do Windows Update",
            "updates",
            "Windows Update",
            "Renomeia caches e restaura serviços originalmente ativos.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Reconstrói caches do Windows Update.",
            600,
        ),
        lambda host, p: deps.updates.reset_components(host),
        validator=_wrap_three_arg(validate_windows_update_reset),
    )
