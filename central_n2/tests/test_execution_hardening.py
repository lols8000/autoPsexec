from __future__ import annotations

import json
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
    RecoveryResult,
    RetryPolicy,
    RiskLevel,
)
from execution.validators import postcheck_succeeded
from remediation import ValidationResult, ValidationStatus
from storage.database import CentralDatabase


def _context(
    *,
    transport: str = "winrm",
    capabilities: dict | None = None,
    ready: bool = True,
) -> ExecutionPolicyContext:
    values = {
        "IsAdmin": True,
        "IsSystem": False,
    }
    if capabilities:
        values.update(capabilities)
    return ExecutionPolicyContext(
        host="PC01",
        ready=ready,
        transport=transport,
        capabilities=values,
    )


def _spec(
    key: str = "test.action",
    **kwargs,
) -> ExecutionAction:
    defaults = {
        "title": "Teste",
        "category": "test",
        "category_label": "Teste",
        "description": "Ação de teste.",
        "operation_class": OperationClass.LIGHT_WRITE,
        "risk": RiskLevel.LOW,
        "impact": "Nenhum.",
    }
    defaults.update(kwargs)
    return ExecutionAction(key=key, **defaults)


def test_policy_blocks_unsupported_transport_before_handler_runs():
    called = 0
    registry = ActionRegistry()

    def handler(host, parameters):
        nonlocal called
        called += 1
        return CommandResult(True, "test", host)

    registry.register(
        BoundExecutionAction(
            spec=_spec(
                allowed_transports=("winrm",),
            ),
            handler=handler,
        )
    )

    engine = ExecutionEngine(registry)
    plan = engine.plan(
        "PC01",
        "test.action",
        {},
        context=_context(transport="psexec"),
    )

    assert plan.allowed is False
    assert any(
        item.key == "transport"
        and item.state.value == "FAIL"
        for item in plan.policy.checks
    )

    try:
        engine.execute(
            "PC01",
            "test.action",
            {},
            context=_context(transport="psexec"),
        )
    except ExecutionBlockedError:
        pass
    else:
        raise AssertionError("Transporte incompatível deveria bloquear a ação")

    assert called == 0


def test_policy_blocks_missing_required_capability():
    registry = ActionRegistry()
    registry.register(
        BoundExecutionAction(
            spec=_spec(
                required_capabilities=("Winget",),
            ),
            handler=lambda host, params: CommandResult(
                True,
                "test",
                host,
            ),
        )
    )

    plan = ExecutionEngine(registry).plan(
        "PC01",
        "test.action",
        {},
        context=_context(capabilities={"Winget": False}),
    )

    assert plan.allowed is False
    assert any(
        item.key == "capability.Winget"
        and item.state.value == "FAIL"
        for item in plan.policy.checks
    )


def test_temporary_disconnect_recovery_turns_indeterminate_into_validated_pass():
    registry = ActionRegistry()

    def handler(host, parameters):
        return CommandResult.failure(
            host,
            "restart-adapter",
            "connection dropped",
            return_code=124,
            transport="winrm",
        ).mark_indeterminate("Conexão caiu durante a mutação.")

    registry.register(
        BoundExecutionAction(
            spec=_spec(
                key="network.restart",
                disconnect_mode=DisconnectMode.TEMPORARY,
                may_break_connectivity=True,
                recovery_timeout_seconds=30,
            ),
            handler=handler,
            after_probe=lambda host, params: CommandResult(
                True,
                "postcheck",
                host,
                data={"Adapter": "Up"},
                transport="winrm",
            ),
            validator=postcheck_succeeded,
        )
    )

    record = ExecutionEngine(registry).execute(
        "PC01",
        "network.restart",
        {},
        context=_context(),
        reconnect=lambda host, timeout, delay: RecoveryResult(
            attempted=True,
            ready=True,
            attempts=2,
            elapsed_seconds=4.2,
            transport="winrm",
            state="READY_WINRM",
        ),
    )

    assert record.remediation.command_result.indeterminate is True
    assert record.recovery is not None
    assert record.recovery.ready is True
    assert record.remediation.validation.status is ValidationStatus.PASS


def test_temporary_disconnect_without_recovery_stays_unknown():
    registry = ActionRegistry()

    registry.register(
        BoundExecutionAction(
            spec=_spec(
                key="network.restart",
                disconnect_mode=DisconnectMode.TEMPORARY,
                may_break_connectivity=True,
                recovery_timeout_seconds=30,
            ),
            handler=lambda host, params: CommandResult.failure(
                host,
                "restart-adapter",
                "connection dropped",
                return_code=124,
                transport="winrm",
            ).mark_indeterminate("Conexão caiu durante a mutação."),
            after_probe=lambda host, params: {
                "_probe_success": False,
            },
            validator=postcheck_succeeded,
        )
    )

    record = ExecutionEngine(registry).execute(
        "PC01",
        "network.restart",
        {},
        context=_context(),
        reconnect=lambda host, timeout, delay: RecoveryResult(
            attempted=True,
            ready=False,
            attempts=5,
            elapsed_seconds=30.0,
            error="timeout",
        ),
    )

    assert record.remediation.validation.status is ValidationStatus.UNKNOWN
    assert "não foi revalidada" in record.remediation.validation.message


def test_pre_execution_transport_failure_is_retried_but_indeterminate_is_not():
    registry = ActionRegistry()
    calls = 0

    def handler(host, parameters):
        nonlocal calls
        calls += 1
        if calls == 1:
            result = CommandResult.failure(
                host,
                "test",
                "WinRM unavailable before dispatch",
                transport="winrm",
            )
            result.metadata["transport_failure_kind"] = "pre_execution"
            return result
        return CommandResult(
            True,
            "test",
            host,
            transport="psexec",
        )

    registry.register(
        BoundExecutionAction(
            spec=_spec(
                retry_policy=RetryPolicy.PRE_EXECUTION_ONLY,
                retry_attempts=3,
                retry_delay_seconds=0,
            ),
            handler=handler,
        )
    )

    record = ExecutionEngine(registry).execute(
        "PC01",
        "test.action",
        {},
        context=_context(),
    )

    assert calls == 2
    assert record.remediation.command_result.success is True
    assert (
        record.remediation.command_result.metadata["execution_attempts"]
        == 2
    )


def test_rollback_restores_original_state_and_validates():
    registry = ActionRegistry()
    state = {"value": "before"}

    def probe(host, parameters):
        return {"value": state["value"]}

    def mutate(host, parameters):
        state["value"] = "after"
        return CommandResult(True, "mutate", host)

    def validate(before, command, after, parameters):
        return ValidationResult(
            ValidationStatus.PASS
            if after["value"] == "after"
            else ValidationStatus.FAIL,
            "mutated",
            after,
        )

    def rollback(host, parameters, before):
        state["value"] = before["value"]
        return CommandResult(True, "rollback", host)

    def validate_rollback(before, command, after, parameters):
        return ValidationResult(
            ValidationStatus.PASS
            if after["value"] == before["value"]
            else ValidationStatus.FAIL,
            "restored",
            after,
        )

    registry.register(
        BoundExecutionAction(
            spec=_spec(
                rollback_strategy="Restaurar estado original.",
            ),
            handler=mutate,
            before_probe=probe,
            after_probe=probe,
            validator=validate,
            rollback_handler=rollback,
            rollback_validator=validate_rollback,
        )
    )

    engine = ExecutionEngine(registry)
    record = engine.execute(
        "PC01",
        "test.action",
        {},
        context=_context(),
    )
    rollback_record = engine.rollback(
        record,
        context=_context(),
    )

    assert record.rollback_available is True
    assert state["value"] == "before"
    assert rollback_record.validation.status is ValidationStatus.PASS


def test_audit_payload_redacts_sensitive_values_and_raw_secrets():
    registry = ActionRegistry()
    secret = "very-secret-value"

    registry.register(
        BoundExecutionAction(
            spec=_spec(
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
                "test",
                host,
                stderr="",
                data={
                    "token": secret,
                    "nested": {"password": secret},
                },
                transport="winrm",
            ),
            before_probe=lambda host, params: {
                "api_key": secret,
            },
            after_probe=lambda host, params: {
                "authorization": secret,
            },
        )
    )

    record = ExecutionEngine(registry).execute(
        "PC01",
        "test.action",
        {"token": secret},
        context=_context(),
        operator="DOMAIN\\operator",
    )

    encoded = json.dumps(
        record.audit_payload(),
        ensure_ascii=False,
        default=str,
    )

    assert secret not in encoded
    assert record.audit_payload()["parameters"]["token"] == "***"


def test_sqlite_v4_persists_audit_columns_and_rollback_link(tmp_path: Path):
    database = CentralDatabase(tmp_path / "central.db")

    registry = ActionRegistry()
    state = {"value": "before"}

    def probe(host, parameters):
        return {"value": state["value"]}

    def rollback(host, parameters, before):
        state["value"] = before["value"]
        return CommandResult(True, "rollback", host, transport="winrm")

    registry.register(
        BoundExecutionAction(
            spec=_spec(
                rollback_strategy="restore",
            ),
            handler=lambda host, params: (
                state.__setitem__("value", "after")
                or CommandResult(
                    True,
                    "mutate",
                    host,
                    transport="winrm",
                )
            ),
            before_probe=probe,
            after_probe=probe,
            rollback_handler=rollback,
        )
    )

    engine = ExecutionEngine(registry)
    record = engine.execute(
        "PC01",
        "test.action",
        {},
        context=_context(),
        operator="operator01",
    )
    execution_id = database.save_execution_record(
        record,
        correlation_id="CORR01",
    )

    rollback_record = engine.rollback(
        record,
        context=_context(),
        operator="operator01",
    )
    rollback_id = database.save_rollback_record(
        rollback_record,
        original_execution_id=execution_id,
        operator="operator01",
        correlation_id="CORR01",
    )

    rows = database.recent_executions("PC01", limit=10)
    original = next(item for item in rows if item["id"] == execution_id)
    rollback_row = next(item for item in rows if item["id"] == rollback_id)

    assert original["operator"] == "operator01"
    assert original["transport"] == "winrm"
    assert original["action_version"] == 1
    assert original["rollback_available"] == 0
    assert rollback_row["is_rollback"] == 1
    assert rollback_row["rollback_of"] == execution_id
