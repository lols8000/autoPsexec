from __future__ import annotations

from pathlib import Path

from core.jobs import OperationClass
from core.result import CommandResult
from execution import (
    ActionRegistry,
    BoundExecutionAction,
    ExecutionAction,
    ExecutionDependencies,
    ExecutionEngine,
    ExecutionParameter,
    ParameterKind,
    RiskLevel,
    build_execution_registry,
    command_completed,
)
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
from remediation import ValidationStatus


class DummyExecutor:
    logger = None


def _registry(tmp_path: Path) -> ActionRegistry:
    executor = DummyExecutor()
    settings = {
        "software": {
            "chrome": {
                "name": "Google Chrome",
                "winget_id": "Google.Chrome",
            }
        },
        "packages": {
            "corp": {
                "source": r"\\server\share\corp.msi",
                "type": "msi",
            }
        },
        "certificates": {
            "ca": {
                "source": r"\\server\share\ca.cer",
                "store": "Root",
            }
        },
        "registry_actions": {
            "policy": {
                "path": r"HKLM:\SOFTWARE\Empresa",
                "name": "Enabled",
                "type": "DWord",
                "value": 1,
            }
        },
    }
    settings_path = tmp_path / "settings.json"
    settings_path.write_text("{}", encoding="utf-8")

    deps = ExecutionDependencies(
        system=SystemModule(executor),
        network=NetworkModule(executor),
        software=SoftwareModule(
            executor,
            settings_path,
            settings=settings,
        ),
        printers=PrintersModule(executor),
        devices=DevicesModule(executor),
        domain=DomainModule(executor),
        users=UsersProfilesModule(executor),
        disk=DiskModule(executor),
        glpi=GLPIModule(
            executor,
            settings_path,
            settings=settings,
        ),
        security=SecurityModule(executor),
        updates=UpdatesModule(executor),
        repair=RepairModule(executor),
        packages=PackagesModule(executor, settings),
        certificates=CertificatesModule(executor, settings),
        registry_actions=RegistryActionsModule(executor, settings),
        files=FileOperationsModule(executor),
    )
    return build_execution_registry(deps)


def test_execution_catalog_is_comprehensive_and_consistent(tmp_path: Path):
    registry = _registry(tmp_path)
    actions = registry.all()

    assert len(actions) >= 59

    keys = [item.spec.key for item in actions]
    assert len(keys) == len(set(keys))

    categories = {item.spec.category for item in actions}
    assert {
        "processes",
        "services",
        "software",
        "network",
        "printers",
        "devices",
        "windows",
        "updates",
        "domain",
        "users",
        "disk",
        "glpi",
        "security",
        "energy",
        "packages",
        "certificates",
        "registry",
        "files",
    }.issubset(categories)

    for bound in actions:
        spec = bound.spec
        assert spec.operation_class is not OperationClass.READ_ONLY
        assert spec.timeout_seconds > 0
        assert spec.requires_confirmation is True
        assert spec.key
        assert spec.title
        assert spec.description
        assert spec.impact

        parameter_keys = [item.key for item in spec.parameters]
        assert len(parameter_keys) == len(set(parameter_keys))

        if spec.risk is RiskLevel.CRITICAL or spec.destructive:
            assert spec.requires_confirmation is True


def test_execution_registry_filters_categories(tmp_path: Path):
    registry = _registry(tmp_path)

    network = registry.by_category("network")
    assert len(network) >= 9
    assert all(item.spec.category == "network" for item in network)

    labels = dict(registry.categories())
    assert labels["network"] == "Rede"
    assert labels["services"] == "Serviços"


def test_execution_engine_coerces_parameters_and_validates():
    registry = ActionRegistry()
    seen = {}

    spec = ExecutionAction(
        key="test.action",
        title="Teste",
        category="test",
        category_label="Teste",
        description="Ação de teste.",
        operation_class=OperationClass.LIGHT_WRITE,
        risk=RiskLevel.LOW,
        impact="Nenhum.",
        timeout_seconds=30,
        parameters=(
            ExecutionParameter(
                "count",
                "Quantidade",
                ParameterKind.INTEGER,
                min_value=1,
                max_value=10,
            ),
            ExecutionParameter(
                "mode",
                "Modo",
                ParameterKind.CHOICE,
                choices=("A", "B"),
            ),
        ),
    )

    def handler(host, parameters):
        seen.update(parameters)
        return CommandResult(
            True,
            "test",
            host,
            data={"ok": True},
        )

    registry.register(
        BoundExecutionAction(
            spec=spec,
            handler=handler,
            validator=command_completed,
        )
    )

    result = ExecutionEngine(registry).execute(
        "PC01",
        "test.action",
        {"count": "4", "mode": "B"},
    )

    assert seen == {"count": 4, "mode": "B"}
    assert result.parameters == seen
    assert result.remediation.validation.status is ValidationStatus.PASS


def test_execution_engine_blocks_unknown_parameters():
    registry = ActionRegistry()
    registry.register(
        BoundExecutionAction(
            spec=ExecutionAction(
                key="test.action",
                title="Teste",
                category="test",
                category_label="Teste",
                description="Teste.",
                operation_class=OperationClass.LIGHT_WRITE,
                risk=RiskLevel.LOW,
                impact="Nenhum.",
                parameters=(),
            ),
            handler=lambda host, params: CommandResult(
                True,
                "test",
                host,
            ),
        )
    )

    try:
        ExecutionEngine(registry).execute(
            "PC01",
            "test.action",
            {"unexpected": "x"},
        )
    except ValueError as exc:
        assert "não reconhecido" in str(exc)
    else:
        raise AssertionError("Parâmetro desconhecido deveria falhar")


def test_indeterminate_execution_becomes_unknown_validation():
    registry = ActionRegistry()

    def handler(host, parameters):
        return CommandResult.failure(
            host,
            "mutation",
            "transport lost",
            return_code=124,
            transport="winrm",
        ).mark_indeterminate(
            "A ação pode ter sido entregue."
        )

    registry.register(
        BoundExecutionAction(
            spec=ExecutionAction(
                key="test.indeterminate",
                title="Teste indeterminado",
                category="test",
                category_label="Teste",
                description="Teste.",
                operation_class=OperationClass.HEAVY_WRITE,
                risk=RiskLevel.HIGH,
                impact="Teste.",
            ),
            handler=handler,
            validator=command_completed,
        )
    )

    result = ExecutionEngine(registry).execute(
        "PC01",
        "test.indeterminate",
    )

    assert result.remediation.command_result.indeterminate is True
    assert (
        result.remediation.validation.status
        is ValidationStatus.UNKNOWN
    )


def test_sensitive_parameters_are_redacted():
    registry = ActionRegistry()

    spec = ExecutionAction(
        key="test.secret",
        title="Teste",
        category="test",
        category_label="Teste",
        description="Teste.",
        operation_class=OperationClass.LIGHT_WRITE,
        risk=RiskLevel.LOW,
        impact="Nenhum.",
        parameters=(
            ExecutionParameter(
                "token",
                "Token",
                sensitive=True,
            ),
        ),
    )
    registry.register(
        BoundExecutionAction(
            spec=spec,
            handler=lambda host, params: CommandResult(
                True,
                "test",
                host,
            ),
        )
    )

    record = ExecutionEngine(registry).execute(
        "PC01",
        "test.secret",
        {"token": "super-secret"},
    )

    assert record.parameters["token"] == "super-secret"
    assert record.public_parameters["token"] == "***"
