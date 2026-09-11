from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

from core.result import CommandResult
from remediation import ValidationResult

from .models import ExecutionAction


Handler = Callable[[str, dict[str, Any]], CommandResult]
Probe = Callable[[str, dict[str, Any]], Any]
Validator = Callable[
    [Any, CommandResult, Any, dict[str, Any]],
    ValidationResult,
]


@dataclass(frozen=True, slots=True)
class BoundExecutionAction:
    spec: ExecutionAction
    handler: Handler
    before_probe: Probe | None = None
    after_probe: Probe | None = None
    validator: Validator | None = None


class ActionRegistry:
    def __init__(self) -> None:
        self._actions: dict[str, BoundExecutionAction] = {}

    def register(self, action: BoundExecutionAction) -> None:
        key = action.spec.key
        if key in self._actions:
            raise ValueError(f"Ação duplicada no catálogo: {key}")
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

    def __len__(self) -> int:
        return len(self._actions)
