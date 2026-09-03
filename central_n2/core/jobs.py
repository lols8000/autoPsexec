from __future__ import annotations

import itertools
import sys
import time
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeout
from dataclasses import dataclass
from typing import Callable, TypeVar

T = TypeVar("T")


@dataclass(slots=True)
class JobStatus:
    label: str
    elapsed_seconds: float
    completed: bool
    timed_out: bool = False


class ResponsiveJobRunner:
    """Executa operações bloqueantes em worker thread e mantém feedback visual no console."""

    def __init__(self, *, heartbeat_seconds: float = 0.2) -> None:
        self.heartbeat_seconds = max(0.05, float(heartbeat_seconds))
        self._pool = ThreadPoolExecutor(max_workers=4, thread_name_prefix="central-n2")

    def run(
        self,
        label: str,
        func: Callable[[], T],
        *,
        timeout: float | None = None,
        on_tick: Callable[[JobStatus], None] | None = None,
    ) -> T:
        started = time.monotonic()
        future = self._pool.submit(func)
        spinner = itertools.cycle("|/-\\")

        while True:
            elapsed = time.monotonic() - started
            if timeout is not None:
                remaining = timeout - elapsed
                if remaining <= 0:
                    future.cancel()
                    self._clear_status_line()
                    if on_tick:
                        on_tick(JobStatus(label, elapsed, False, True))
                    raise TimeoutError(f"Operação '{label}' excedeu {timeout:.0f}s")
                wait_for = min(self.heartbeat_seconds, remaining)
            else:
                wait_for = self.heartbeat_seconds

            try:
                result = future.result(timeout=wait_for)
                elapsed = time.monotonic() - started
                self._clear_status_line()
                if on_tick:
                    on_tick(JobStatus(label, elapsed, True, False))
                return result
            except FutureTimeout:
                elapsed = time.monotonic() - started
                if timeout is not None and elapsed >= timeout:
                    future.cancel()
                    self._clear_status_line()
                    if on_tick:
                        on_tick(JobStatus(label, elapsed, False, True))
                    raise TimeoutError(f"Operação '{label}' excedeu {timeout:.0f}s")
                if on_tick:
                    on_tick(JobStatus(label, elapsed, False, False))
                else:
                    sys.stdout.write(f"\r{next(spinner)} {label}... {elapsed:5.1f}s")
                    sys.stdout.flush()

    @staticmethod
    def _clear_status_line() -> None:
        sys.stdout.write("\r" + " " * 100 + "\r")
        sys.stdout.flush()

    def shutdown(self) -> None:
        self._pool.shutdown(wait=False, cancel_futures=True)
