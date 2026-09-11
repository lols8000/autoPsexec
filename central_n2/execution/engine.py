from __future__ import annotations

from typing import Any

from remediation import RemediationEngine, RemediationSpec

from .models import ExecutionRecord
from .registry import ActionRegistry, BoundExecutionAction, Probe, Validator
from .validators import command_completed





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
    ) -> None:
        self.registry = registry
        self.remediation_engine = remediation_engine or RemediationEngine()

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

    def execute(
        self,
        host: str,
        action_key: str,
        parameters: dict[str, Any] | None = None,
    ) -> ExecutionRecord:
        bound = self.registry.get(action_key)
        parsed = self.validate_parameters(bound, parameters)

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

        remediation = self.remediation_engine.execute(
            host,
            RemediationSpec(
                key=bound.spec.key,
                title=bound.spec.title,
                impact=bound.spec.impact,
                requires_confirmation=bound.spec.requires_confirmation,
                requires_reboot=bound.spec.requires_reboot,
                disruptive=bound.spec.may_break_connectivity,
                rollback=None,
            ),
            lambda target: bound.handler(target, parsed),
            before_probe=before_probe,
            after_probe=after_probe,
            validator=validator,
        )

        return ExecutionRecord(
            action=bound.spec,
            host=host,
            parameters=parsed,
            remediation=remediation,
        )
