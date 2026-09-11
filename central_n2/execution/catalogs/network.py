from __future__ import annotations

import time

from core.jobs import OperationClass
from ..models import ExecutionAction, ExecutionParameter, ParameterKind, RiskLevel
from ..validators import command_completed
from .common import ExecutionDependencies, _adapter_state, _register


def register(registry, deps: ExecutionDependencies) -> None:
    _register(
        registry,
        ExecutionAction(
            "network.flush_dns",
            "Limpar cache DNS",
            "network",
            "Rede",
            "Limpa o cache DNS local.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Baixo impacto; consultas DNS serão refeitas.",
            120,
        ),
        lambda host, p: deps.network.flush_dns(host),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "network.register_dns",
            "Registrar DNS",
            "network",
            "Rede",
            "Solicita novo registro DNS dos adaptadores.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Baixo impacto.",
            180,
        ),
        lambda host, p: deps.network.register_dns(host),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "network.renew_dhcp",
            "Release/Renew DHCP",
            "network",
            "Rede",
            "Libera e renova concessões DHCP.",
            OperationClass.DISRUPTIVE,
            RiskLevel.HIGH,
            "Pode interromper a conectividade da estação.",
            300,
            may_break_connectivity=True,
        ),
        lambda host, p: deps.network.renew_dhcp(host),
        validator=command_completed,
    )
    for key, title, network_handler, reboot in (
        (
            "network.reset_winsock",
            "Resetar Winsock",
            deps.network.reset_winsock,
            True,
        ),
        (
            "network.reset_tcpip",
            "Resetar TCP/IP",
            deps.network.reset_tcpip,
            True,
        ),
        (
            "network.clear_arp",
            "Limpar tabela ARP",
            deps.network.clear_arp,
            False,
        ),
    ):
        _register(
            registry,
            ExecutionAction(
                key,
                title,
                "network",
                "Rede",
                title,
                OperationClass.DISRUPTIVE if reboot else OperationClass.LIGHT_WRITE,
                RiskLevel.HIGH if reboot else RiskLevel.MEDIUM,
                (
                    "Pode exigir reinicialização e afetar conectividade."
                    if reboot
                    else "Conexões locais podem precisar reaprender vizinhos."
                ),
                300,
                requires_reboot=reboot,
                may_break_connectivity=reboot,
            ),
            lambda host, p, fn=network_handler: fn(host),
            validator=command_completed,
        )

    adapter_parameter = ExecutionParameter(
        "adapter_name",
        "Nome do adaptador",
        ParameterKind.TEXT,
    )
    for key, title, enabled, risk in (
        (
            "network.adapter_enable",
            "Habilitar adaptador",
            True,
            RiskLevel.MEDIUM,
        ),
        (
            "network.adapter_disable",
            "Desabilitar adaptador",
            False,
            RiskLevel.CRITICAL,
        ),
    ):
        _register(
            registry,
            ExecutionAction(
                key,
                title,
                "network",
                "Rede",
                f"{title} pelo nome exato.",
                OperationClass.DISRUPTIVE,
                risk,
                "Pode derrubar o transporte remoto se for o adaptador em uso.",
                180,
                may_break_connectivity=True,
                destructive=not enabled,
                parameters=(adapter_parameter,),
            ),
            lambda host, p, state=enabled: deps.network.set_adapter_state(
                host,
                p["adapter_name"],
                enabled=state,
            ),
            before_probe=lambda host, p: deps.network.adapter_status(
                host,
                p["adapter_name"],
            ),
            after_probe=lambda host, p: deps.network.adapter_status(
                host,
                p["adapter_name"],
            ),
            validator=_adapter_state(enabled),
        )

    def restart_adapter_probe(host: str, p: dict[str, Any]):
        time.sleep(8)
        return deps.network.adapter_status(host, p["adapter_name"])

    _register(
        registry,
        ExecutionAction(
            "network.adapter_restart",
            "Reiniciar adaptador",
            "network",
            "Rede",
            "Agenda disable/enable local para evitar abandonar o adaptador desabilitado.",
            OperationClass.DISRUPTIVE,
            RiskLevel.HIGH,
            "A conexão pode cair por alguns segundos.",
            240,
            may_break_connectivity=True,
            parameters=(adapter_parameter,),
        ),
        lambda host, p: deps.network.restart_adapter(
            host,
            p["adapter_name"],
        ),
        before_probe=lambda host, p: deps.network.adapter_status(
            host,
            p["adapter_name"],
        ),
        after_probe=restart_adapter_probe,
        validator=_adapter_state(True),
    )
