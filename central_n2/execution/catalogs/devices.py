from __future__ import annotations

from core.jobs import OperationClass
from ..models import ExecutionAction, ExecutionParameter, ParameterKind, RiskLevel
from ..validators import command_completed
from .common import ExecutionDependencies, _device_state, _register


def register(registry, deps: ExecutionDependencies) -> None:
    instance = ExecutionParameter(
        "instance_id",
        "PNP InstanceId",
        ParameterKind.TEXT,
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
                    ParameterKind.TEXT,
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
                    ParameterKind.TEXT,
                ),
            ),
        ),
        lambda host, p: deps.devices.remove_driver_package(
            host,
            p["inf_name"],
        ),
        validator=command_completed,
    )
