from __future__ import annotations

from dataclasses import replace
from datetime import datetime, timezone
from time import perf_counter
from typing import Any, Callable

from core.result import CommandResult
from remediation import (
    RemediationEngine,
    RemediationSpec,
    ValidationResult,
    ValidationStatus,
)

from .models import (
    DisconnectMode,
    ExecutionPlan,
    ExecutionRecord,
    ExecutionRollbackRecord,
    RecoveryResult,
)
from .policy import (
    ExecutionPolicy,
    ExecutionPolicyContext,
)
from .registry import (
    ActionRegistry,
    BoundExecutionAction,
    Probe,
    Validator,
)
from .validators import command_completed


ReconnectCallback = Callable[[str, int, int], RecoveryResult]


class ExecutionBlockedError(RuntimeError):
    pass


def _bind_probe(
    probe: Probe,
    parameters: dict[str, Any],
):
    def bound(target: str) -> Any:
        return probe(target, parameters)

    return bound


def _bind_validator(
    validator: Validator,
    parameters: dict[str, Any],
):
    def bound(before, command, after):
        return validator(before, command, after, parameters)

    return bound


class ExecutionEngine:
    def __init__(
        self,
        registry: ActionRegistry,
        *,
        remediation_engine: RemediationEngine | None = None,
        policy: ExecutionPolicy | None = None,
    ) -> None:
        self.registry = registry
        self.remediation_engine = remediation_engine or RemediationEngine()
        self.policy = policy or ExecutionPolicy()

    @staticmethod
    def validate_parameters(
        action: BoundExecutionAction,
        raw_parameters: dict[str, Any] | None,
    ) -> dict[str, Any]:
        raw = raw_parameters or {}
        known = {parameter.key for parameter in action.spec.parameters}
        unknown = set(raw) - known
        if unknown:
            names = ", ".join(sorted(unknown))
            raise ValueError(
                f"Parâmetro(s) não reconhecido(s): {names}"
            )

        return {
            parameter.key: parameter.parse(raw.get(parameter.key))
            for parameter in action.spec.parameters
        }

    def plan(
        self,
        host: str,
        action_key: str,
        parameters: dict[str, Any] | None,
        *,
        context: ExecutionPolicyContext,
    ) -> ExecutionPlan:
        bound = self.registry.get(action_key)
        parsed = self.validate_parameters(bound, parameters)
        report = self.policy.evaluate(
            bound.spec,
            context,
            parsed,
            custom_preconditions=bound.preconditions,
        )
        return ExecutionPlan(
            action=bound.spec,
            host=host,
            parameters=parsed,
            policy=report,
        )

    @staticmethod
    def _blocked_message(plan: ExecutionPlan) -> str:
        failures = getattr(plan.policy, "failures", [])
        details = "; ".join(
            str(item.message)
            for item in failures
        )
        return details or "A política de execução bloqueou a ação."

    def execute(
        self,
        host: str,
        action_key: str,
        parameters: dict[str, Any] | None = None,
        *,
        context: ExecutionPolicyContext,
        operator: str | None = None,
        reconnect: ReconnectCallback | None = None,
    ) -> ExecutionRecord:
        bound = self.registry.get(action_key)
        plan = self.plan(
            host,
            action_key,
            parameters,
            context=context,
        )
        if not plan.allowed:
            raise ExecutionBlockedError(
                self._blocked_message(plan)
            )

        parsed = plan.parameters
        before_probe = (
            _bind_probe(bound.before_probe, parsed)
            if bound.before_probe is not None
            else None
        )
        after_probe = (
            _bind_probe(bound.after_probe, parsed)
            if bound.after_probe is not None
            else None
        )
        validator = _bind_validator(
            bound.validator or command_completed,
            parsed,
        )

        recovery: RecoveryResult | None = None

        def action(target: str) -> CommandResult:
            nonlocal recovery
            result = bound.handler(target, parsed)
            result.metadata["execution_action"] = bound.spec.key
            result.metadata["action_version"] = bound.spec.action_version
            result.metadata["disconnect_mode"] = (
                bound.spec.disconnect_mode.value
            )

            if bound.spec.disconnect_mode is DisconnectMode.TEMPORARY:
                if reconnect is None:
                    recovery = RecoveryResult(
                        attempted=False,
                        ready=False,
                        error="Callback de reconexão não configurado.",
                    )
                else:
                    recovery = reconnect(
                        target,
                        bound.spec.recovery_timeout_seconds,
                        bound.spec.recovery_delay_seconds,
                    )
                result.metadata["recovery"] = {
                    "attempted": recovery.attempted,
                    "ready": recovery.ready,
                    "attempts": recovery.attempts,
                    "elapsed_seconds": recovery.elapsed_seconds,
                    "transport": recovery.transport,
                    "state": recovery.state,
                    "error": recovery.error,
                }
            return result

        def validate(
            before: Any,
            command: CommandResult,
            after: Any,
        ) -> ValidationResult:
            if bound.spec.disconnect_mode is DisconnectMode.TEMPORARY:
                if recovery is None or not recovery.ready:
                    return ValidationResult(
                        ValidationStatus.UNKNOWN,
                        (
                            "A ação esperava uma desconexão temporária, "
                            "mas a estação não foi revalidada dentro do prazo."
                        ),
                        {
                            "recovery": recovery,
                            "after": after,
                        },
                    )

                effective = command
                if command.indeterminate:
                    metadata = dict(command.metadata)
                    metadata.pop("indeterminate", None)
                    metadata.pop("indeterminate_reason", None)
                    effective = replace(
                        command,
                        success=True,
                        metadata=metadata,
                    )
                return validator(before, effective, after)

            return validator(before, command, after)

        started_wall = datetime.now(timezone.utc).isoformat()
        started = perf_counter()

        remediation = self.remediation_engine.execute(
            host,
            RemediationSpec(
                key=bound.spec.key,
                title=bound.spec.title,
                impact=bound.spec.impact,
                requires_confirmation=bound.spec.requires_confirmation,
                requires_reboot=bound.spec.requires_reboot,
                disruptive=bound.spec.may_break_connectivity,
                rollback=bound.spec.rollback_strategy,
            ),
            action,
            before_probe=before_probe,
            after_probe=after_probe,
            validator=validate,
        )

        duration_ms = int((perf_counter() - started) * 1000)
        finished_wall = datetime.now(timezone.utc).isoformat()

        return ExecutionRecord(
            action=bound.spec,
            host=host,
            parameters=parsed,
            remediation=remediation,
            policy=plan.policy,
            started_at=started_wall,
            finished_at=finished_wall,
            duration_ms=duration_ms,
            operator=operator,
            recovery=recovery,
            rollback_available=bound.rollback_handler is not None,
        )

    def rollback(
        self,
        record: ExecutionRecord,
        *,
        context: ExecutionPolicyContext,
        operator: str | None = None,
    ) -> ExecutionRollbackRecord:
        bound = self.registry.get(record.action.key)
        if bound.rollback_handler is None:
            raise ExecutionBlockedError(
                "Esta ação não possui rollback automático."
            )

        report = self.policy.evaluate(
            bound.spec,
            context,
            record.parameters,
            custom_preconditions=bound.rollback_preconditions,
        )
        if not report.allowed:
            failures = "; ".join(
                item.message
                for item in report.failures
            )
            raise ExecutionBlockedError(
                failures or "Rollback bloqueado pela política."
            )

        started_wall = datetime.now(timezone.utc).isoformat()
        started = perf_counter()

        result = bound.rollback_handler(
            record.host,
            record.parameters,
            record.remediation.before,
        )

        after = None
        if bound.after_probe is not None:
            try:
                after = bound.after_probe(
                    record.host,
                    record.parameters,
                )
            except Exception as exc:
                after = {
                    "probe_error": (
                        f"{type(exc).__name__}: {exc}"
                    )
                }

        validator = bound.rollback_validator
        if validator is None:
            validation = command_completed(
                record.remediation.before,
                result,
                after,
                record.parameters,
            )
        else:
            validation = validator(
                record.remediation.before,
                result,
                after,
                record.parameters,
            )

        finished_wall = datetime.now(timezone.utc).isoformat()
        duration_ms = int((perf_counter() - started) * 1000)

        record.rollback_result = result

        return ExecutionRollbackRecord(
            action_key=record.action.key,
            host=record.host,
            result=result,
            validation=validation,
            started_at=started_wall,
            finished_at=finished_wall,
            duration_ms=duration_ms,
        )
