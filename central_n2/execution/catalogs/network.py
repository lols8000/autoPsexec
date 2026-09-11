from __future__ import annotations

from typing import Any

from core.jobs import OperationClass
from core.result import CommandResult
from remediation import ValidationResult, ValidationStatus

from ..models import (
    DisconnectMode,
    ExecutionAction,
    ExecutionParameter,
    ParameterKind,
    RiskLevel,
    SelectorKind,
)
from ..validators import (
    command_completed,
    postcheck_succeeded,
)
from .common import ExecutionDependencies, _adapter_state, _register


def _payload(value: Any) -> dict[str, Any]:
    if isinstance(value, CommandResult):
        return value.data if isinstance(value.data, dict) else {}
    return value if isinstance(value, dict) else {}


def _rollback_adapter(
    deps: ExecutionDependencies,
    host: str,
    parameters: dict[str, Any],
    before: Any,
) -> CommandResult:
    original = _payload(before)
    status = str(original.get("Status") or "").casefold()
    if not status:
        return CommandResult.failure(
            host,
            "network.adapter.rollback",
            "Estado original do adaptador não está disponível.",
        )
    enabled = status != "disabled"
    return deps.network.set_adapter_state(
        host,
        parameters["adapter_name"],
        enabled=enabled,
    )


def _validate_adapter_rollback(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    original = _payload(before)
    current = _payload(after)
    expected = str(original.get("Status") or "").casefold()
    actual = str(current.get("Status") or "").casefold()

    if command.indeterminate:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "Rollback do adaptador teve resultado indeterminado.",
            current,
        )
    if expected and actual == expected:
        return ValidationResult(
            ValidationStatus.PASS,
            f"Estado original do adaptador restaurado: {original.get('Status')}.",
            current,
        )
    return ValidationResult(
        ValidationStatus.FAIL if not command.success else ValidationStatus.UNKNOWN,
        "Estado original do adaptador não foi confirmado.",
        current,
    )


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
            idempotent=True,
            tags=("rede", "dns", "cache"),
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
            idempotent=True,
            tags=("rede", "dns", "register"),
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
            "Libera e renova concessões DHCP e revalida a estação.",
            OperationClass.DISRUPTIVE,
            RiskLevel.HIGH,
            "Pode interromper a conectividade da estação por alguns segundos.",
            600,
            may_break_connectivity=True,
            disconnect_mode=DisconnectMode.TEMPORARY,
            recovery_timeout_seconds=180,
            recovery_delay_seconds=5,
            tags=("rede", "dhcp", "ip", "renew"),
        ),
        lambda host, p: deps.network.renew_dhcp(host),
        after_probe=lambda host, p: deps.network.ip_configuration(host),
        validator=postcheck_succeeded,
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
                idempotent=True,
                tags=("rede", "tcpip" if "tcpip" in key else "winsock" if "winsock" in key else "arp"),
            ),
            lambda host, p, fn=network_handler: fn(host),
            validator=command_completed,
        )

    adapter_parameter = ExecutionParameter(
        "adapter_name",
        "Adaptador",
        ParameterKind.TEXT,
        selector=SelectorKind.ADAPTER,
        help_text="Selecione um adaptador inventariado ou informe o nome exato.",
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
                240,
                may_break_connectivity=True,
                destructive=not enabled,
                parameters=(adapter_parameter,),
                idempotent=True,
                required_capabilities=("NetAdapter",),
                disconnect_mode=(
                    DisconnectMode.NONE
                    if enabled
                    else DisconnectMode.TERMINAL
                ),
                rollback_strategy="Restaurar o estado Up/Disabled observado antes da ação.",
                tags=("rede", "adaptador", "nic", "enable" if enabled else "disable"),
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
            rollback_handler=lambda host, p, before, d=deps: _rollback_adapter(
                d,
                host,
                p,
                before,
            ),
            rollback_validator=_validate_adapter_rollback,
        )

    _register(
        registry,
        ExecutionAction(
            "network.adapter_restart",
            "Reiniciar adaptador",
            "network",
            "Rede",
            "Agenda disable/enable local e revalida a estação após a queda esperada.",
            OperationClass.DISRUPTIVE,
            RiskLevel.HIGH,
            "A conexão pode cair por alguns segundos.",
            360,
            may_break_connectivity=True,
            parameters=(adapter_parameter,),
            idempotent=True,
            required_capabilities=("NetAdapter",),
            disconnect_mode=DisconnectMode.TEMPORARY,
            recovery_timeout_seconds=180,
            recovery_delay_seconds=8,
            tags=("rede", "adaptador", "nic", "restart"),
        ),
        lambda host, p: deps.network.restart_adapter(
            host,
            p["adapter_name"],
        ),
        before_probe=lambda host, p: deps.network.adapter_status(
            host,
            p["adapter_name"],
        ),
        after_probe=lambda host, p: deps.network.adapter_status(
            host,
            p["adapter_name"],
        ),
        validator=_adapter_state(True),
    )
