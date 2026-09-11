from __future__ import annotations

from dataclasses import dataclass
from typing import Any

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
from remediation import ValidationResult, ValidationStatus

from ..models import ExecutionAction
from ..registry import ActionRegistry, BoundExecutionAction


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


def _process_absent(before, command, after, parameters):
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


def _startup_type(before, command, after, parameters):
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
    def validate(before, command, after, parameters):
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
    def validate(before, command, after, parameters):
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
            return ValidationResult(ValidationStatus.PASS, message, data)
        return _unknown_or_fail(
            command,
            "O comando concluiu, mas o estado PnP final não pôde ser confirmado.",
            data,
        )
    return validate


def _secure_channel(before, command, after, parameters):
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


def _glpi_running(before, command, after, parameters):
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


def _profile_removed(before, command, after, parameters):
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
    def validate(before, command, after, parameters):
        data = _data(after)
        if data.get("Exists") is expected:
            return ValidationResult(
                ValidationStatus.PASS,
                "Impressora confirmada." if expected
                else "Impressora removida e ausência confirmada.",
                data,
            )
        return _unknown_or_fail(
            command,
            "Estado final da impressora não foi confirmado.",
            data,
        )
    return validate


def _defender_ready(before, command, after, parameters):
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


def _update_install(before, command, after, parameters):
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
        before, command, after
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
