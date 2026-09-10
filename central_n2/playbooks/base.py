from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable

from core.result import CommandResult


@dataclass(frozen=True, slots=True)
class PlaybookStep:
    key: str
    label: str


@dataclass(frozen=True, slots=True)
class PlaybookSpec:
    key: str
    title: str
    steps: tuple[PlaybookStep, ...]


@dataclass(slots=True)
class PlaybookExecution:
    playbook: str
    host: str
    started_at: str
    finished_at: str
    steps: list[dict[str, Any]] = field(default_factory=list)


StepCallback = Callable[[int, int, PlaybookStep], None]
Collector = Callable[[str], CommandResult]


class PlaybookRunner:
    def run(
        self,
        spec: PlaybookSpec,
        host: str,
        collectors: dict[str, Collector],
        *,
        on_step: StepCallback | None = None,
    ) -> PlaybookExecution:
        started = datetime.now().astimezone().isoformat(
            timespec="seconds"
        )
        results: list[dict[str, Any]] = []
        total = len(spec.steps)

        for index, step in enumerate(spec.steps, 1):
            if on_step:
                on_step(index, total, step)

            collector = collectors.get(step.key)
            if collector is None:
                results.append(
                    {
                        "key": step.key,
                        "label": step.label,
                        "success": False,
                        "error": "Coletor não configurado",
                    }
                )
                continue

            try:
                result = collector(host)
                results.append(
                    {
                        "key": step.key,
                        "label": step.label,
                        "success": result.success,
                        "transport": result.transport,
                        "data": result.data,
                        "stdout": result.stdout,
                        "error": result.stderr,
                        "duration_ms": result.duration_ms,
                        "indeterminate": result.indeterminate,
                    }
                )
            except Exception as exc:
                results.append(
                    {
                        "key": step.key,
                        "label": step.label,
                        "success": False,
                        "error": f"{type(exc).__name__}: {exc}",
                    }
                )

        return PlaybookExecution(
            playbook=spec.key,
            host=host,
            started_at=started,
            finished_at=datetime.now().astimezone().isoformat(
                timespec="seconds"
            ),
            steps=results,
        )
