from __future__ import annotations

from typing import Any

from remediation import RemediationEngine, RemediationSpec

from .models import ExecutionRecord
from .registry import ActionRegistry, BoundExecutionAction
from .validators import command_completed


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
            before_probe=(
                None
                if bound.before_probe is None
                else lambda target: bound.before_probe(target, parsed)
            ),
            after_probe=(
                None
                if bound.after_probe is None
                else lambda target: bound.after_probe(target, parsed)
            ),
            validator=lambda before, command, after: (
                bound.validator or command_completed
            )(before, command, after, parsed),
        )

        return ExecutionRecord(
            action=bound.spec,
            host=host,
            parameters=parsed,
            remediation=remediation,
        )
