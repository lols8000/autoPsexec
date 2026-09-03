from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Callable

from core.result import CommandResult


@dataclass(slots=True)
class RetryPolicy:
    max_attempts: int = 2
    base_delay_seconds: float = 0.5

    @staticmethod
    def retryable(result: CommandResult) -> bool:
        if result.success:
            return False
        text = (result.stderr or "").lower()
        no_retry = ("access is denied", "acesso negado", "nxDOMAIN", "não existe", "not found")
        if any(item.lower() in text for item in no_retry):
            return False
        retry_markers = ("timeout", "timed out", "connection reset", "cannotconnect", "wsman", "temporariamente")
        return result.return_code == 124 or any(item in text for item in retry_markers)

    def run(self, action: Callable[[], CommandResult], *, on_retry: Callable[[int, CommandResult], None] | None = None) -> CommandResult:
        last: CommandResult | None = None
        for attempt in range(1, max(1, self.max_attempts) + 1):
            last = action()
            if last.success or not self.retryable(last) or attempt >= self.max_attempts:
                return last
            if on_retry:
                on_retry(attempt + 1, last)
            time.sleep(self.base_delay_seconds * (2 ** (attempt - 1)))
        assert last is not None
        return last
