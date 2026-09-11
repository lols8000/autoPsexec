from __future__ import annotations

from pathlib import Path

from core.jobs import OperationClass
from core.result import CommandResult
from execution import (
    ActionRegistry,
    BoundExecutionAction,
    DisconnectMode,
    ExecutionAction,
    ExecutionBlockedError,
    ExecutionEngine,
    ExecutionParameter,
    ExecutionPolicyContext,
    ParameterKind,
    PrivilegeLevel,
    RecoveryResult,
    RiskLevel,
)
from execution.config_validation import validate_execution_configuration
from execution.validators import (
    command_completed,
    postcheck_succeeded,
)
from remediation import ValidationResult, ValidationStatus
from storage.database import CentralDatabase


def _context(
    *,
    transport: str = "winrm",
    capabilities: dict | None = None,
) -> ExecutionPolicyContext:
    values = {
        "IsAdmin": True,
        "IsSystem": False,
    }
    if capabilities:
        values.update(capabilities)
    return ExecutionPolicyContext(
        host="PC01",
        ready=True,
        transport=transport,
        capabilities=values,
    )


def test_policy_blocks_transport_before_handler():
    called = False

    def handler(host, parameters):
        nonlocal called
        called = True
        return CommandResult(True, "x", host)

    registry = ActionRegistry()
    registry.register(
        BoundExecutionAction(
            spec=ExecutionAction(
                key="software.test",
                title="Software",
                category="software",
                category_label="Software",
                description="Teste.",
                operation_class=OperationClass.HEAVY_WRITE,
                risk=RiskLevel.MEDIUM,
                impact="Teste.",
                allowed_transports=("local", "winrm"),
                required_capabilities=("Winget",),
                required_privilege=PrivilegeLevel.USER_CONTEXT,
            ),
            handler=handler,
        )
    )

    engine = ExecutionEngine(registry)
    plan = engine.plan(
        "PC01",
        "software.test",
        {},
        context=_context(
            transport="psexec",
            capabilities={
                "Winget": False,
                "IsSystem": True,
            },
        ),
    )

    assert plan.allowed is False
    assert len(plan.policy.failures) >= 2

    try:
        engine.execute(
            "PC01",
            "software.test",
            {},
            context=_context(
                transport="psexec",
                capabilities={
                    "Winget": False,
                    "IsSystem": True,
                },
            ),
        )
    except ExecutionBlockedError:
        pass
    else:
        raise AssertionError("Ação incompatível deveria ser bloqueada")

    assert called is False


def test_policy_allows_required_capability_and_user_context():
    registry = ActionRegistry()
    registry.register(
        BoundExecutionAction(
            spec=ExecutionAction(
                key="software.test",
                title="Software",
                category="software",
                category_label="Software",
                description="Teste.",
                operation_class=OperationClass.HEAVY_WRITE,
                risk=RiskLevel.MEDIUM,
                impact="Teste.",
                allowed_transports=("winrm",),
                required_capabilities=("Winget",),
                required_privilege=PrivilegeLevel.USER_CONTEXT,
            ),
            handler=lambda host, params: CommandResult(
                True,
                "x",
                host,
            ),
        )
    )

    plan = ExecutionEngine(registry).plan(
        "PC01",
        "software.test",
        {},
        context=_context(
            capabilities={"Winget": True},
        ),
    )

    assert plan.allowed is True
    assert not plan.policy.failures


def test_expected_disconnect_recovery_can_validate_final_state():
    registry = ActionRegistry()
    command = CommandResult.failure(
        "PC01",
        "renew",
        "transport lost",
        return_code=124,
        transport="winrm",
    ).mark_indeterminate("Conexão caiu durante ação esperada.")

    registry.register(
        BoundExecutionAction(
            spec=ExecutionAction(
                key="network.test",
                title="Rede",
                category="network",
                category_label="Rede",
                description="Teste de recovery.",
                operation_class=OperationClass.DISRUPTIVE,
                risk=RiskLevel.HIGH,
                impact="Conectividade.",
                disconnect_mode=DisconnectMode.TEMPORARY,
                recovery_timeout_seconds=30,
            ),
            handler=lambda host, params: command,
            after_probe=lambda host, params: CommandResult(
                True,
                "probe",
                host,
                data={"IPv4": "10.0.0.2"},
                transport="winrm",
            ),
            validator=postcheck_succeeded,
        )
    )

    result = ExecutionEngine(registry).execute(
        "PC01",
        "network.test",
        {},
        context=_context(),
        reconnect=lambda host, timeout, delay: RecoveryResult(
            attempted=True,
            ready=True,
            attempts=2,
            elapsed_seconds=4.0,
            transport="winrm",
            state="READY_WINRM",
        ),
    )

    assert result.remediation.command_result.indeterminate is True
    assert result.recovery is not None
    assert result.recovery.ready is True
    assert result.remediation.validation.status is ValidationStatus.PASS


def test_expected_disconnect_without_recovery_is_unknown():
    registry = ActionRegistry()
    registry.register(
        BoundExecutionAction(
            spec=ExecutionAction(
                key="network.test",
                title="Rede",
                category="network",
                category_label="Rede",
                description="Teste de recovery.",
                operation_class=OperationClass.DISRUPTIVE,
                risk=RiskLevel.HIGH,
                impact="Conectividade.",
                disconnect_mode=DisconnectMode.TEMPORARY,
                recovery_timeout_seconds=30,
            ),
            handler=lambda host, params: CommandResult(
                True,
                "renew",
                host,
                transport="winrm",
            ),
            after_probe=lambda host, params: CommandResult(
                False,
                "probe",
                host,
                stderr="offline",
            ),
            validator=postcheck_succeeded,
        )
    )

    result = ExecutionEngine(registry).execute(
        "PC01",
        "network.test",
        {},
        context=_context(),
        reconnect=lambda host, timeout, delay: RecoveryResult(
            attempted=True,
            ready=False,
            attempts=5,
            elapsed_seconds=30,
            error="timeout",
        ),
    )

    assert result.remediation.validation.status is ValidationStatus.UNKNOWN


def test_execution_rollback_restores_original_state():
    state = {"Status": "Stopped"}

    def probe(host, params):
        return dict(state)

    def start(host, params):
        state["Status"] = "Running"
        return CommandResult(True, "start", host)

    def rollback(host, params, before):
        state["Status"] = before["Status"]
        return CommandResult(True, "rollback", host)

    def validate_running(before, command, after, params):
        return ValidationResult(
            ValidationStatus.PASS
            if after["Status"] == "Running"
            else ValidationStatus.FAIL,
            "running",
            after,
        )

    def validate_original(before, command, after, params):
        return ValidationResult(
            ValidationStatus.PASS
            if after["Status"] == before["Status"]
            else ValidationStatus.FAIL,
            "restored",
            after,
        )

    registry = ActionRegistry()
    registry.register(
        BoundExecutionAction(
            spec=ExecutionAction(
                key="service.test",
                title="Serviço",
                category="services",
                category_label="Serviços",
                description="Teste.",
                operation_class=OperationClass.LIGHT_WRITE,
                risk=RiskLevel.MEDIUM,
                impact="Teste.",
                rollback_strategy="Restaurar estado.",
            ),
            handler=start,
            before_probe=probe,
            after_probe=probe,
            validator=validate_running,
            rollback_handler=rollback,
            rollback_validator=validate_original,
        )
    )
    engine = ExecutionEngine(registry)

    record = engine.execute(
        "PC01",
        "service.test",
        {},
        context=_context(),
    )
    assert record.rollback_available is True
    assert state["Status"] == "Running"

    reverted = engine.rollback(
        record,
        context=_context(),
    )

    assert state["Status"] == "Stopped"
    assert reverted.validation.status is ValidationStatus.PASS


def test_action_search_uses_tags_and_description():
    registry = ActionRegistry()
    registry.register(
        BoundExecutionAction(
            spec=ExecutionAction(
                key="printer.spooler",
                title="Reiniciar Spooler",
                category="printers",
                category_label="Impressão",
                description="Reinicia serviço de impressão.",
                operation_class=OperationClass.LIGHT_WRITE,
                risk=RiskLevel.LOW,
                impact="Baixo.",
                tags=("fila", "spooler", "impressora"),
            ),
            handler=lambda host, params: CommandResult(True, "x", host),
        )
    )

    assert registry.search("fila")[0].spec.key == "printer.spooler"
    assert registry.search("impressão")[0].spec.key == "printer.spooler"
    assert registry.search("inexistente") == []


def test_catalog_validation_disables_only_invalid_entries():
    report = validate_execution_configuration(
        {
            "packages": {
                "good": {
                    "source": r"\\server\share\app.msi",
                    "type": "msi",
                },
                "bad": {
                    "source": r"\\server\share\app.ps1",
                    "type": "ps1",
                },
            },
            "certificates": {
                "bad": {
                    "source": r"\\server\share\private.pfx",
                    "store": "My",
                }
            },
            "registry_actions": {
                "good": {
                    "path": r"HKLM:\SOFTWARE\Empresa",
                    "name": "Enabled",
                    "mode": "set",
                    "type": "DWord",
                    "value": 1,
                },
                "bad": {
                    "path": r"HKCU:\SOFTWARE\Empresa",
                    "name": "Enabled",
                    "mode": "set",
                    "type": "DWord",
                    "value": 1,
                },
            },
        }
    )

    assert set(report.settings["packages"]) == {"good"}
    assert report.settings["certificates"] == {}
    assert set(report.settings["registry_actions"]) == {"good"}
    assert len(report.issues) == 3


def test_execution_audit_payload_redacts_sensitive_parameters():
    registry = ActionRegistry()
    registry.register(
        BoundExecutionAction(
            spec=ExecutionAction(
                key="secret.test",
                title="Secret",
                category="test",
                category_label="Test",
                description="Test.",
                operation_class=OperationClass.LIGHT_WRITE,
                risk=RiskLevel.LOW,
                impact="None.",
                parameters=(
                    ExecutionParameter(
                        "token",
                        "Token",
                        ParameterKind.TEXT,
                        sensitive=True,
                    ),
                ),
            ),
            handler=lambda host, params: CommandResult(
                True,
                "secret command",
                host,
                data={"token": params["token"]},
            ),
            validator=command_completed,
        )
    )

    record = ExecutionEngine(registry).execute(
        "PC01",
        "secret.test",
        {"token": "super-secret-token"},
        context=_context(),
    )
    payload = record.audit_payload()

    assert payload["parameters"]["token"] == "***"
    assert payload["command"]["data"]["token"] == "***"
    assert "super-secret-token" not in str(payload)


def test_database_schema_v4_persists_rich_execution(tmp_path: Path):
    database = CentralDatabase(tmp_path / "central.db")

    registry = ActionRegistry()
    registry.register(
        BoundExecutionAction(
            spec=ExecutionAction(
                key="test.persist",
                title="Persist",
                category="test",
                category_label="Test",
                description="Test.",
                operation_class=OperationClass.LIGHT_WRITE,
                risk=RiskLevel.LOW,
                impact="None.",
            ),
            handler=lambda host, params: CommandResult(
                True,
                "x",
                host,
                transport="winrm",
            ),
        )
    )
    record = ExecutionEngine(registry).execute(
        "PC01",
        "test.persist",
        {},
        context=_context(),
        operator="DOMAIN\\operator",
    )

    execution_id = database.save_execution_record(
        record,
        correlation_id="CORR001",
    )
    rows = database.recent_executions("PC01", limit=1)

    assert database.SCHEMA_VERSION == 4
    assert execution_id == rows[0]["id"]
    assert rows[0]["operator"] == "DOMAIN\\operator"
    assert rows[0]["transport"] == "winrm"
    assert rows[0]["risk"] == "LOW"
    assert rows[0]["action_version"] == 1
    assert rows[0]["correlation_id"] == "CORR001"
