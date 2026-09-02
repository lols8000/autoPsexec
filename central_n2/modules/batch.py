from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict
from typing import Callable, Iterable, List, Dict, Any

from core.result import CommandResult


class BatchRunner:
    def __init__(self, max_workers: int = 5) -> None:
        self.max_workers = max(1, int(max_workers))

    def run(
        self,
        hosts: Iterable[str],
        action: Callable[[str], CommandResult],
    ) -> List[Dict[str, Any]]:
        targets = [h.strip() for h in hosts if h and h.strip()]
        results: List[Dict[str, Any]] = []
        with ThreadPoolExecutor(max_workers=min(self.max_workers, max(1, len(targets)))) as pool:
            futures = {pool.submit(action, host): host for host in targets}
            for future in as_completed(futures):
                host = futures[future]
                try:
                    result = future.result()
                except Exception as exc:
                    result = CommandResult.failure(host, "batch_action", str(exc))
                results.append(asdict(result))
        return sorted(results, key=lambda r: r["host"].lower())
