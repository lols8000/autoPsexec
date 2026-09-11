from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

from core.result import CommandResult
from remediation import ValidationResult

from .models import DisconnectMode, ExecutionAction, RetryPolicy
from .policy import CustomPrecondition


Handler = Callable[[str, dict[str, Any]], CommandResult]
Probe = Callable[[str, dict[str, Any]], Any]
Validator = Callable[
    [Any, CommandResult, Any, dict[str, Any]],
    ValidationResult,
]
RollbackHandler = Callable[
    [str, dict[str, Any], Any],
    CommandResult,
]


@dataclass(frozen=True, slots=True)
class BoundExecutionAction:
    spec: ExecutionAction
    handler: Handler
    before_probe: Probe | None = None
    after_probe: Probe | None = None
    validator: Validator | None = None
    preconditions: tuple[CustomPrecondition, ...] = ()
    rollback_preconditions: tuple[CustomPrecondition, ...] = ()
    rollback_handler: RollbackHandler | None = None
    rollback_validator: Validator | None = None


class ActionRegistry:
    def __init__(self) -> None:
        self._actions: dict[str, BoundExecutionAction] = {}

    @staticmethod
    def _validate_contract(action: BoundExecutionAction) -> None:
        spec = action.spec

        if spec.timeout_seconds <= 0:
            raise ValueError(
                f"Ação {spec.key}: timeout_seconds deve ser > 0."
            )
        if not spec.allowed_transports:
            raise ValueError(
                f"Ação {spec.key}: informe ao menos um transporte."
            )

        allowed = {"local", "winrm", "psexec"}
        invalid = {
            item.casefold()
            for item in spec.allowed_transports
        } - allowed
        if invalid:
            raise ValueError(
                f"Ação {spec.key}: transporte(s) inválido(s): "
                + ", ".join(sorted(invalid))
            )

        if spec.destructive and not spec.requires_confirmation:
            raise ValueError(
                f"Ação {spec.key}: ação destrutiva exige confirmação."
            )

        if (
            spec.retry_policy is RetryPolicy.SAFE_TRANSIENT
            and not spec.idempotent
        ):
            raise ValueError(
                f"Ação {spec.key}: SAFE_TRANSIENT exige idempotência."
            )

        if (
            spec.disconnect_mode is DisconnectMode.TEMPORARY
            and spec.recovery_timeout_seconds <= 0
        ):
            raise ValueError(
                f"Ação {spec.key}: recovery_timeout_seconds deve ser > 0."
            )

        if (
            action.rollback_handler is not None
            and not spec.rollback_strategy
        ):
            raise ValueError(
                f"Ação {spec.key}: rollback handler exige estratégia documentada."
            )

    def register(self, action: BoundExecutionAction) -> None:
        key = action.spec.key
        if key in self._actions:
            raise ValueError(f"Ação duplicada no catálogo: {key}")
        self._validate_contract(action)
        self._actions[key] = action

    def get(self, key: str) -> BoundExecutionAction:
        try:
            return self._actions[key]
        except KeyError as exc:
            raise KeyError(f"Ação não registrada: {key}") from exc

    def all(self) -> list[BoundExecutionAction]:
        return sorted(
            self._actions.values(),
            key=lambda item: (
                item.spec.category_label.casefold(),
                item.spec.title.casefold(),
            ),
        )

    def categories(self) -> list[tuple[str, str]]:
        values = {
            item.spec.category: item.spec.category_label
            for item in self._actions.values()
        }
        return sorted(
            values.items(),
            key=lambda item: item[1].casefold(),
        )

    def by_category(self, category: str) -> list[BoundExecutionAction]:
        return [
            item
            for item in self.all()
            if item.spec.category == category
        ]

    def search(self, query: str) -> list[BoundExecutionAction]:
        normalized = query.strip().casefold()
        if not normalized:
            return self.all()
        return [
            item
            for item in self.all()
            if normalized
            in " ".join(
                (
                    item.spec.key,
                    item.spec.title,
                    item.spec.category_label,
                    item.spec.description,
                    " ".join(item.spec.tags),
                )
            ).casefold()
        ]

    def __len__(self) -> int:
        return len(self._actions)
