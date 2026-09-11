from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Any

from core.jobs import OperationClass
from core.result import CommandResult
from modules.certificates import CertificatesModule
from modules.devices import DevicesModule
from modules.disk import DiskModule
from modules.domain import DomainModule
from modules.file_ops import FileOperationsModule
from modules.glpi import GLPIModule
from modules.network import NetworkModule
from modules.packages import PackagesModule
from modules.printers import PrintersModule
from modules.registry_actions import RegistryActionsModule
from modules.repair import RepairModule
from modules.security import SecurityModule
from modules.software import SoftwareModule
from modules.system import SystemModule
from modules.updates import UpdatesModule
from modules.users_profiles import UsersProfilesModule
from remediation import (
    ValidationResult,
    ValidationStatus,
    validate_cleanup,
    validate_gpupdate,
    validate_spooler,
    validate_windows_update_reset,
)

from .models import (
    ExecutionAction,
    ExecutionParameter,
    ParameterKind,
    RiskLevel,
)
from .registry import ActionRegistry, BoundExecutionAction
from .validators import (
    command_completed,
    command_field_true,
    service_running,
    service_stopped,
)


@dataclass(slots=True)
class ExecutionDependencies:
    system: SystemModule
    network: NetworkModule
    software: SoftwareModule
    printers: PrintersModule
    devices: DevicesModule
    domain: DomainModule
    users: UsersProfilesModule
    disk: DiskModule
    glpi: GLPIModule
    security: SecurityModule
    updates: UpdatesModule
    repair: RepairModule
    packages: PackagesModule
    certificates: CertificatesModule
    registry_actions: RegistryActionsModule
    files: FileOperationsModule


def _data(value: Any) -> dict[str, Any]:
    if isinstance(value, CommandResult):
        return value.data if isinstance(value.data, dict) else {}
    return value if isinstance(value, dict) else {}


def _unknown_or_fail(
    command: CommandResult,
    message: str,
    evidence: Any,
) -> ValidationResult:
    if command.indeterminate:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "A execução teve resultado indeterminado.",
            evidence,
        )
    return ValidationResult(
        ValidationStatus.UNKNOWN if command.success else ValidationStatus.FAIL,
        message if command.success else command.stderr or message,
        evidence,
    )


def _process_absent(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    data = _data(after)
    if "Exists" not in data:
        return _unknown_or_fail(
            command,
            "Não foi possível confirmar o encerramento do processo.",
            data,
        )
    if data.get("Exists") is False:
        return ValidationResult(
            ValidationStatus.PASS,
            "Processo não está mais em execução.",
            data,
        )
    return ValidationResult(
        ValidationStatus.FAIL,
        "Processo continua em execução.",
        data,
    )


def _startup_type(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    data = _data(after)
    expected = str(parameters.get("startup_type") or "")
    actual = str(data.get("StartType") or "")
    if actual and actual.casefold() == expected.casefold():
        return ValidationResult(
            ValidationStatus.PASS,
            f"StartType confirmado como {expected}.",
            data,
        )
    return _unknown_or_fail(
        command,
        f"StartType não foi confirmado como {expected}.",
        data,
    )


def _adapter_state(expected_enabled: bool):
    def validate(
        before: Any,
        command: CommandResult,
        after: Any,
        parameters: dict[str, Any],
    ) -> ValidationResult:
        data = _data(after)
        status = str(data.get("Status") or "").casefold()
        if not status:
            return _unknown_or_fail(
                command,
                "Não foi possível confirmar o estado do adaptador.",
                data,
            )
        if expected_enabled and status != "disabled":
            return ValidationResult(
                ValidationStatus.PASS,
                f"Adaptador habilitado; estado atual: {data.get('Status')}.",
                data,
            )
        if not expected_enabled and status == "disabled":
            return ValidationResult(
                ValidationStatus.PASS,
                "Adaptador confirmado como Disabled.",
                data,
            )
        return ValidationResult(
            ValidationStatus.FAIL,
            f"Estado inesperado do adaptador: {data.get('Status')}.",
            data,
        )

    return validate


def _device_state(expected_enabled: bool):
    def validate(
        before: Any,
        command: CommandResult,
        after: Any,
        parameters: dict[str, Any],
    ) -> ValidationResult:
        data = _data(after)
        if not data.get("Exists"):
            return _unknown_or_fail(
                command,
                "Dispositivo não pôde ser consultado após a ação.",
                data,
            )
        status = str(data.get("Status") or "").casefold()
        problem = str(data.get("Problem") or "").casefold()
        if expected_enabled:
            ok = status == "ok" or problem in {"0", "cm_prob_none", "none"}
            message = "Dispositivo confirmado como habilitado."
        else:
            ok = "disabled" in problem or problem in {"22", "cm_prob_disabled"}
            message = "Dispositivo confirmado como desabilitado."
        if ok:
            return ValidationResult(
                ValidationStatus.PASS,
                message,
                data,
            )
        return _unknown_or_fail(
            command,
            "O comando concluiu, mas o estado PnP final não pôde ser confirmado.",
            data,
        )

    return validate


def _secure_channel(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    data = _data(after)
    if data.get("SecureChannel") is True:
        return ValidationResult(
            ValidationStatus.PASS,
            "Secure channel confirmado como íntegro.",
            data,
        )
    return _unknown_or_fail(
        command,
        "Secure channel não foi confirmado após o reparo.",
        data,
    )


def _glpi_running(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    data = _data(after)
    if data.get("Installed") is True and data.get("Running") is True:
        return ValidationResult(
            ValidationStatus.PASS,
            "GLPI Agent instalado e em execução.",
            data,
        )
    return _unknown_or_fail(
        command,
        "GLPI Agent não foi confirmado como Running.",
        data,
    )


def _profile_removed(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    data = _data(after)
    if data.get("Exists") is False:
        return ValidationResult(
            ValidationStatus.PASS,
            "Perfil não está mais registrado no Win32_UserProfile.",
            data,
        )
    return _unknown_or_fail(
        command,
        "Perfil ainda existe ou não pôde ser validado.",
        data,
    )


def _printer_exists(expected: bool):
    def validate(
        before: Any,
        command: CommandResult,
        after: Any,
        parameters: dict[str, Any],
    ) -> ValidationResult:
        data = _data(after)
        if data.get("Exists") is expected:
            return ValidationResult(
                ValidationStatus.PASS,
                (
                    "Impressora confirmada."
                    if expected
                    else "Impressora removida e ausência confirmada."
                ),
                data,
            )
        return _unknown_or_fail(
            command,
            "Estado final da impressora não foi confirmado.",
            data,
        )

    return validate


def _defender_ready(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    data = _data(after)
    if data.get("AntivirusEnabled") is True:
        return ValidationResult(
            ValidationStatus.PASS,
            "Microsoft Defender respondeu e permanece habilitado.",
            data,
        )
    return _unknown_or_fail(
        command,
        "Não foi possível confirmar o Defender após a ação.",
        data,
    )


def _update_install(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    data = _data(command)
    if command.indeterminate:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "Instalação pode ter avançado; valide o Windows Update antes de repetir.",
            data,
        )
    if command.success and data.get("ResultCode") in {0, 2, 3}:
        return ValidationResult(
            ValidationStatus.PASS,
            (
                f"Windows Update concluiu. Instaladas: "
                f"{data.get('Installed', 0)}; "
                f"reboot: {'SIM' if data.get('RebootRequired') else 'NÃO'}."
            ),
            data,
        )
    return ValidationResult(
        ValidationStatus.FAIL,
        command.stderr or "Windows Update não concluiu com resultado aceito.",
        data,
    )


def _wrap_three_arg(validator):
    return lambda before, command, after, parameters: validator(
        before,
        command,
        after,
    )


def _register(
    registry: ActionRegistry,
    spec: ExecutionAction,
    handler,
    *,
    before_probe=None,
    after_probe=None,
    validator=None,
) -> None:
    registry.register(
        BoundExecutionAction(
            spec=spec,
            handler=handler,
            before_probe=before_probe,
            after_probe=after_probe,
            validator=validator,
        )
    )


def build_execution_registry(
    deps: ExecutionDependencies,
) -> ActionRegistry:
    registry = ActionRegistry()

    text = ParameterKind.TEXT
    integer = ParameterKind.INTEGER
    choice = ParameterKind.CHOICE

    # PROCESSOS
    _register(
        registry,
        ExecutionAction(
            "process.kill_name",
            "Finalizar processo por nome",
            "processes",
            "Processos",
            "Finaliza todas as instâncias do processo informado.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.MEDIUM,
            "Pode fechar aplicação e causar perda de trabalho não salvo.",
            120,
            destructive=True,
            parameters=(
                ExecutionParameter("process_name", "Nome do processo", text),
            ),
        ),
        lambda host, p: deps.system.kill_process(host, p["process_name"]),
        before_probe=lambda host, p: deps.system.process_status(
            host,
            process_name=p["process_name"],
        ),
        after_probe=lambda host, p: deps.system.process_status(
            host,
            process_name=p["process_name"],
        ),
        validator=_process_absent,
    )
    _register(
        registry,
        ExecutionAction(
            "process.kill_pid",
            "Finalizar processo por PID",
            "processes",
            "Processos",
            "Finaliza uma instância específica pelo PID.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.MEDIUM,
            "Pode encerrar aplicação ou componente do sistema.",
            120,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "pid",
                    "PID",
                    integer,
                    min_value=1,
                    max_value=2_147_483_647,
                ),
            ),
        ),
        lambda host, p: deps.system.kill_process_pid(host, p["pid"]),
        before_probe=lambda host, p: deps.system.process_status(
            host,
            pid=p["pid"],
        ),
        after_probe=lambda host, p: deps.system.process_status(
            host,
            pid=p["pid"],
        ),
        validator=_process_absent,
    )

    # SERVIÇOS
    service_name = (
        ExecutionParameter("service_name", "Nome do serviço", text),
    )
    for key, title, method, validator in (
        ("service.start", "Iniciar serviço", "start", service_running),
        ("service.stop", "Parar serviço", "stop", service_stopped),
        ("service.restart", "Reiniciar serviço", "restart", service_running),
    ):
        _register(
            registry,
            ExecutionAction(
                key,
                title,
                "services",
                "Serviços",
                f"{title} e valida o estado final.",
                OperationClass.LIGHT_WRITE,
                RiskLevel.MEDIUM,
                "Pode afetar aplicações dependentes do serviço.",
                180,
                destructive=method == "stop",
                parameters=service_name,
            ),
            lambda host, p, action=method: deps.system.service_action(
                host,
                p["service_name"],
                action,
            ),
            before_probe=lambda host, p: deps.system.service_status(
                host,
                p["service_name"],
            ),
            after_probe=lambda host, p: deps.system.service_status(
                host,
                p["service_name"],
            ),
            validator=validator,
        )
    _register(
        registry,
        ExecutionAction(
            "service.startup",
            "Alterar tipo de inicialização",
            "services",
            "Serviços",
            "Altera Automatic/Manual/Disabled para um serviço.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.HIGH,
            "Uma configuração incorreta pode impedir serviço essencial no próximo boot.",
            180,
            parameters=(
                ExecutionParameter("service_name", "Nome do serviço", text),
                ExecutionParameter(
                    "startup_type",
                    "Tipo de inicialização",
                    choice,
                    choices=("Automatic", "Manual", "Disabled"),
                ),
            ),
        ),
        lambda host, p: deps.system.set_service_startup(
            host,
            p["service_name"],
            p["startup_type"],
        ),
        before_probe=lambda host, p: deps.system.service_status(
            host,
            p["service_name"],
        ),
        after_probe=lambda host, p: deps.system.service_status(
            host,
            p["service_name"],
        ),
        validator=_startup_type,
    )

    # SOFTWARE
    software_keys = tuple(sorted(deps.software.catalog))
    if software_keys:
        software_parameter = ExecutionParameter(
            "software_key",
            "Software homologado",
            choice,
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
                ),
                lambda host, p, fn=handler: fn(
                    host,
                    p["software_key"],
                ),
                validator=command_completed,
            )

    # REDE
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
    for key, title, handler, reboot in (
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
            lambda host, p, fn=handler: fn(host),
            validator=command_completed,
        )

    adapter_parameter = ExecutionParameter(
        "adapter_name",
        "Nome do adaptador",
        text,
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

    # IMPRESSÃO
    _register(
        registry,
        ExecutionAction(
            "printer.restart_spooler",
            "Reiniciar Spooler",
            "printers",
            "Impressão",
            "Reinicia Spooler e valida Running.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Jobs em processamento podem ser interrompidos temporariamente.",
            180,
        ),
        lambda host, p: deps.printers.restart_spooler(host),
        before_probe=lambda host, p: deps.printers.spooler_status(host),
        after_probe=lambda host, p: deps.printers.spooler_status(host),
        validator=_wrap_three_arg(validate_spooler),
    )
    _register(
        registry,
        ExecutionAction(
            "printer.clear_queue",
            "Limpar fila de impressão",
            "printers",
            "Impressão",
            "Remove jobs de uma impressora ou de todas quando o nome ficar vazio.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Trabalhos pendentes serão perdidos.",
            180,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "printer_name",
                    "Impressora (vazio = todas)",
                    text,
                    required=False,
                ),
            ),
        ),
        lambda host, p: deps.printers.clear_queue(
            host,
            p.get("printer_name"),
        ),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "printer.add_connection",
            "Adicionar impressora compartilhada",
            "printers",
            "Impressão",
            r"Adiciona conexão \\\\servidor\\fila.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.MEDIUM,
            "Adiciona impressora ao sistema.",
            240,
            parameters=(
                ExecutionParameter(
                    "connection_name",
                    "UNC da impressora",
                    text,
                ),
            ),
        ),
        lambda host, p: deps.printers.add_connection(
            host,
            p["connection_name"],
        ),
        after_probe=lambda host, p: deps.printers.printer_status(
            host,
            p["connection_name"],
        ),
        validator=_printer_exists(True),
    )
    _register(
        registry,
        ExecutionAction(
            "printer.remove",
            "Remover impressora",
            "printers",
            "Impressão",
            "Remove uma impressora pelo nome exato.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Remove configuração de impressão da estação.",
            180,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "printer_name",
                    "Nome da impressora",
                    text,
                ),
            ),
        ),
        lambda host, p: deps.printers.remove_printer(
            host,
            p["printer_name"],
        ),
        before_probe=lambda host, p: deps.printers.printer_status(
            host,
            p["printer_name"],
        ),
        after_probe=lambda host, p: deps.printers.printer_status(
            host,
            p["printer_name"],
        ),
        validator=_printer_exists(False),
    )

    # DISPOSITIVOS / DRIVERS
    instance = ExecutionParameter(
        "instance_id",
        "PNP InstanceId",
        text,
    )
    for key, title, enabled, risk in (
        ("device.enable", "Habilitar dispositivo", True, RiskLevel.MEDIUM),
        ("device.disable", "Desabilitar dispositivo", False, RiskLevel.HIGH),
    ):
        _register(
            registry,
            ExecutionAction(
                key,
                title,
                "devices",
                "Drivers / Dispositivos",
                title,
                OperationClass.HEAVY_WRITE,
                risk,
                "Pode afetar hardware crítico, inclusive rede e armazenamento.",
                240,
                destructive=not enabled,
                may_break_connectivity=not enabled,
                parameters=(instance,),
            ),
            lambda host, p, state=enabled: deps.devices.set_device_state(
                host,
                p["instance_id"],
                enabled=state,
            ),
            before_probe=lambda host, p: deps.devices.device_status(
                host,
                p["instance_id"],
            ),
            after_probe=lambda host, p: deps.devices.device_status(
                host,
                p["instance_id"],
            ),
            validator=_device_state(enabled),
        )
    _register(
        registry,
        ExecutionAction(
            "device.rescan",
            "Reexaminar hardware",
            "devices",
            "Drivers / Dispositivos",
            "Solicita nova enumeração PnP.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Baixo impacto.",
            240,
        ),
        lambda host, p: deps.devices.rescan(host),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "driver.install_inf",
            "Instalar driver por INF",
            "devices",
            "Drivers / Dispositivos",
            "Instala INF já presente na estação usando pnputil.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Pode trocar driver de dispositivo.",
            900,
            requires_reboot=True,
            parameters=(
                ExecutionParameter(
                    "path",
                    "Caminho absoluto do INF",
                    text,
                ),
            ),
        ),
        lambda host, p: deps.devices.install_driver_inf(
            host,
            p["path"],
        ),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "driver.remove_package",
            "Remover pacote de driver",
            "devices",
            "Drivers / Dispositivos",
            "Remove oemNN.inf com /uninstall /force.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.CRITICAL,
            "Pode remover driver em uso e exigir reinstalação/reboot.",
            900,
            destructive=True,
            requires_reboot=True,
            parameters=(
                ExecutionParameter(
                    "inf_name",
                    "Pacote (oemNN.inf)",
                    text,
                ),
            ),
        ),
        lambda host, p: deps.devices.remove_driver_package(
            host,
            p["inf_name"],
        ),
        validator=command_completed,
    )

    # WINDOWS / REPARO
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

    # WINDOWS UPDATE
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

    # DOMÍNIO / GPO
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

    # USUÁRIOS / PERFIS / SESSÕES
    _register(
        registry,
        ExecutionAction(
            "session.logoff",
            "Encerrar sessão de usuário",
            "users",
            "Usuários / Perfis",
            "Executa logoff pelo ID da sessão.",
            OperationClass.DISRUPTIVE,
            RiskLevel.HIGH,
            "Aplicações abertas podem perder trabalho não salvo.",
            180,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "session_id",
                    "ID da sessão",
                    integer,
                    min_value=0,
                    max_value=65535,
                ),
            ),
        ),
        lambda host, p: deps.system.logoff_session(
            host,
            p["session_id"],
        ),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "profile.clean_temp",
            "Limpar TEMP de perfil",
            "users",
            "Usuários / Perfis",
            "Limpa AppData\\Local\\Temp do SID informado.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.MEDIUM,
            "Arquivos temporários em uso são preservados quando bloqueados.",
            600,
            parameters=(
                ExecutionParameter("sid", "SID do perfil", text),
            ),
        ),
        lambda host, p: deps.users.clean_profile_temp(
            host,
            p["sid"],
        ),
        before_probe=lambda host, p: deps.users.profile_status(
            host,
            p["sid"],
        ),
        after_probe=lambda host, p: deps.users.profile_status(
            host,
            p["sid"],
        ),
        validator=_wrap_three_arg(validate_cleanup),
    )
    _register(
        registry,
        ExecutionAction(
            "profile.remove",
            "Remover perfil de usuário",
            "users",
            "Usuários / Perfis",
            "Remove somente perfil não carregado e não especial.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.CRITICAL,
            "Dados locais do perfil podem ser removidos permanentemente.",
            600,
            destructive=True,
            parameters=(
                ExecutionParameter("sid", "SID do perfil", text),
            ),
        ),
        lambda host, p: deps.users.remove_profile(host, p["sid"]),
        before_probe=lambda host, p: deps.users.profile_status(
            host,
            p["sid"],
        ),
        after_probe=lambda host, p: deps.users.profile_status(
            host,
            p["sid"],
        ),
        validator=_profile_removed,
    )

    # DISCO / LIMPEZA
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

    # GLPI
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

    # SEGURANÇA / DEFENDER
    _register(
        registry,
        ExecutionAction(
            "defender.signatures",
            "Atualizar assinaturas do Defender",
            "security",
            "Segurança / Defender",
            "Executa Update-MpSignature.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Baixo impacto; requer Defender disponível.",
            900,
        ),
        lambda host, p: deps.security.update_defender_signatures(host),
        after_probe=lambda host, p: deps.security.defender_status(host),
        validator=_defender_ready,
    )
    for key, title, scan_type, timeout, risk in (
        (
            "defender.quick_scan",
            "Executar verificação rápida",
            "QuickScan",
            2400,
            RiskLevel.LOW,
        ),
        (
            "defender.full_scan",
            "Executar verificação completa",
            "FullScan",
            7800,
            RiskLevel.MEDIUM,
        ),
    ):
        _register(
            registry,
            ExecutionAction(
                key,
                title,
                "security",
                "Segurança / Defender",
                title,
                OperationClass.HEAVY_WRITE,
                risk,
                "Pode consumir CPU/disco durante a varredura.",
                timeout,
            ),
            lambda host, p, st=scan_type: deps.security.defender_scan(
                host,
                st,
            ),
            before_probe=lambda host, p: deps.security.defender_status(host),
            after_probe=lambda host, p: deps.security.defender_status(host),
            validator=_defender_ready,
        )

    # ENERGIA / COMUNICAÇÃO
    _register(
        registry,
        ExecutionAction(
            "energy.restart",
            "Reiniciar estação",
            "energy",
            "Energia / Sessões",
            "Agenda reinicialização administrativa.",
            OperationClass.DISRUPTIVE,
            RiskLevel.CRITICAL,
            "Interrompe sessões e indisponibiliza a estação temporariamente.",
            180,
            may_break_connectivity=True,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "delay_seconds",
                    "Atraso em segundos",
                    integer,
                    required=False,
                    default=0,
                    min_value=0,
                    max_value=3600,
                ),
            ),
        ),
        lambda host, p: deps.system.restart(
            host,
            p["delay_seconds"],
        ),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "energy.shutdown",
            "Desligar estação",
            "energy",
            "Energia / Sessões",
            "Agenda desligamento administrativo.",
            OperationClass.DISRUPTIVE,
            RiskLevel.CRITICAL,
            "Interrompe sessões e desliga a estação.",
            180,
            may_break_connectivity=True,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "delay_seconds",
                    "Atraso em segundos",
                    integer,
                    required=False,
                    default=0,
                    min_value=0,
                    max_value=3600,
                ),
            ),
        ),
        lambda host, p: deps.system.shutdown(
            host,
            p["delay_seconds"],
        ),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "energy.abort",
            "Cancelar shutdown/restart pendente",
            "energy",
            "Energia / Sessões",
            "Executa shutdown /a.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Cancela temporizador de energia pendente.",
            120,
        ),
        lambda host, p: deps.system.abort_shutdown(host),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "session.message",
            "Enviar mensagem para usuários",
            "energy",
            "Energia / Sessões",
            "Envia mensagem para sessões interativas.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Exibe mensagem aos usuários conectados.",
            120,
            parameters=(
                ExecutionParameter(
                    "message",
                    "Mensagem",
                    text,
                ),
            ),
        ),
        lambda host, p: deps.system.send_message(
            host,
            p["message"],
        ),
        validator=command_completed,
    )

    # PACOTES HOMOLOGADOS
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
                        choice,
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

    # CERTIFICADOS
    certificate_keys = deps.certificates.keys()
    if certificate_keys:
        _register(
            registry,
            ExecutionAction(
                "certificate.import",
                "Importar certificado homologado",
                "certificates",
                "Certificados",
                "Importa .cer/.crt permitido em LocalMachine.",
                OperationClass.HEAVY_WRITE,
                RiskLevel.HIGH,
                "Altera stores de confiança do computador.",
                300,
                parameters=(
                    ExecutionParameter(
                        "certificate_key",
                        "Certificado",
                        choice,
                        choices=certificate_keys,
                    ),
                ),
            ),
            lambda host, p: deps.certificates.import_certificate(
                host,
                p["certificate_key"],
            ),
            validator=command_field_true(
                "Imported",
                pass_message="Certificado importado.",
                fail_message="Importação não foi confirmada.",
            ),
        )

    # REGISTRO HOMOLOGADO
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
                        choice,
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

    # ARQUIVOS
    _register(
        registry,
        ExecutionAction(
            "file.ensure_directory",
            "Criar / garantir diretório",
            "files",
            "Arquivos",
            "Cria diretório remoto validado.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Cria estrutura de diretórios.",
            180,
            parameters=(
                ExecutionParameter("path", "Caminho", text),
            ),
        ),
        lambda host, p: deps.files.ensure_directory(
            host,
            p["path"],
        ),
        validator=command_field_true(
            "Exists",
            pass_message="Diretório confirmado.",
            fail_message="Diretório não foi confirmado.",
        ),
    )
    _register(
        registry,
        ExecutionAction(
            "file.move",
            "Mover / renomear caminho",
            "files",
            "Arquivos",
            "Move arquivo ou diretório entre caminhos absolutos.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Muda localização de dados e pode afetar aplicações.",
            600,
            destructive=True,
            parameters=(
                ExecutionParameter("source", "Origem", text),
                ExecutionParameter("destination", "Destino", text),
            ),
        ),
        lambda host, p: deps.files.move_path(
            host,
            p["source"],
            p["destination"],
        ),
        validator=command_field_true(
            "DestinationExists",
            pass_message="Destino confirmado após a movimentação.",
            fail_message="Destino não foi confirmado.",
        ),
    )
    _register(
        registry,
        ExecutionAction(
            "file.remove",
            "Excluir arquivo",
            "files",
            "Arquivos",
            "Exclui somente arquivo; não remove diretório recursivamente.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.CRITICAL,
            "Exclusão permanente de arquivo.",
            300,
            destructive=True,
            parameters=(
                ExecutionParameter("path", "Arquivo", text),
            ),
        ),
        lambda host, p: deps.files.remove_file(
            host,
            p["path"],
        ),
        validator=command_field_true(
            "Removed",
            pass_message="Arquivo removido e ausência confirmada.",
            fail_message="Remoção não foi confirmada.",
        ),
    )

    return registry
