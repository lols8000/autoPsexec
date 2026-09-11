from __future__ import annotations

from core.jobs import OperationClass
from remediation import validate_gpupdate

from ..models import ExecutionAction, RiskLevel
from ..validators import command_field_true, service_running
from .common import ExecutionDependencies, _register, _secure_channel, _wrap_three_arg


def register(registry, deps: ExecutionDependencies) -> None:
    _register(
        registry,
        ExecutionAction(
            "domain.gpupdate",
            "Executar GPUpdate /force",
            "domain",
            "Domínio / GPO",
            "Atualiza políticas de computador/usuário.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Políticas podem alterar configuração da estação.",
            420,
        ),
        lambda host, p: deps.domain.gpupdate(host),
        after_probe=lambda host, p: deps.domain.gpresult(host),
        validator=_wrap_three_arg(validate_gpupdate),
    )
    _register(
        registry,
        ExecutionAction(
            "domain.secure_channel_repair",
            "Reparar secure channel",
            "domain",
            "Domínio / GPO",
            "Executa Test-ComputerSecureChannel -Repair.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Altera relação segura da estação com o domínio.",
            300,
        ),
        lambda host, p: deps.domain.repair_secure_channel(host),
        before_probe=lambda host, p: deps.domain.status(host),
        after_probe=lambda host, p: deps.domain.status(host),
        validator=_secure_channel,
    )
    _register(
        registry,
        ExecutionAction(
            "domain.time_restart",
            "Reiniciar serviço de horário",
            "domain",
            "Domínio / GPO",
            "Reinicia w32time.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Baixo impacto.",
            180,
        ),
        lambda host, p: deps.domain.restart_time_service(host),
        after_probe=lambda host, p: deps.system.service_status(
            host,
            "w32time",
        ),
        validator=service_running,
    )
    _register(
        registry,
        ExecutionAction(
            "domain.time_resync",
            "Ressincronizar horário",
            "domain",
            "Domínio / GPO",
            "Reinicia w32time e executa /resync /rediscover.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.MEDIUM,
            "Pode ajustar o relógio da estação.",
            240,
        ),
        lambda host, p: deps.domain.resync_time(host),
        validator=command_field_true(
            "ResyncSucceeded",
            pass_message="Ressincronização confirmada.",
            fail_message="Ressincronização não foi confirmada.",
        ),
    )
    _register(
        registry,
        ExecutionAction(
            "domain.kerberos_purge_machine",
            "Limpar tickets Kerberos da conta da máquina",
            "domain",
            "Domínio / GPO",
            "Executa klist purge no LogonId SYSTEM 0x3e7.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Tickets serão renovados; pode afetar autenticação momentaneamente.",
            180,
            destructive=True,
        ),
        lambda host, p: deps.domain.purge_system_kerberos(host),
        validator=command_field_true(
            "Purged",
            pass_message="Tickets da conta da máquina foram limpos.",
            fail_message="Purge Kerberos não foi confirmado.",
        ),
    )
