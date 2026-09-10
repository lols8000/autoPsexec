from __future__ import annotations

from concurrent.futures import as_completed
from dataclasses import asdict
from typing import Any, Callable, Iterable

from core.jobs import JobManager, OperationClass
from core.result import CommandResult


class BatchRunner:
    """Executa lotes usando o mesmo JobManager da aplicação."""

    def __init__(
        self,
        max_workers: int = 5,
        *,
        job_manager: JobManager | None = None,
    ) -> None:
        self.max_workers = max(1, int(max_workers))
        self._manager = job_manager or JobManager(
            max_workers=self.max_workers
        )
        self._owns_manager = job_manager is None

    def run(
        self,
        hosts: Iterable[str],
        action: Callable[[str], CommandResult],
        *,
        operation_class: OperationClass = OperationClass.READ_ONLY,
        correlation_id: str | None = None,
    ) -> list[dict[str, Any]]:
        targets = list(
            dict.fromkeys(
                host.strip()
                for host in hosts
                if host and host.strip()
            )
        )
        if not targets:
            return []

        futures = {}
        for host in targets:
            _, future = self._manager.submit(
                host,
                f"Lote: {host}",
                lambda target=host: action(target),
                operation_class=operation_class,
                correlation_id=correlation_id,
            )
            futures[future] = host

        results: list[dict[str, Any]] = []
        for future in as_completed(futures):
            host = futures[future]
            try:
                result = future.result()
            except Exception as exc:
                result = CommandResult.failure(
                    host,
                    "batch_action",
                    f"{type(exc).__name__}: {exc}",
                )
            results.append(asdict(result))

        return sorted(
            results,
            key=lambda item: str(item["host"]).casefold(),
        )

    def shutdown(self) -> None:
        if self._owns_manager:
            self._manager.shutdown()
