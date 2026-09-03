from __future__ import annotations

import itertools
import sys
import threading
import time
import uuid
from concurrent.futures import Future, ThreadPoolExecutor, TimeoutError as FutureTimeout
from contextlib import contextmanager
from dataclasses import dataclass
from enum import Enum
from typing import Callable, TypeVar

T = TypeVar("T")


class JobState(str, Enum):
    QUEUED = "QUEUED"
    RUNNING = "RUNNING"
    SUCCESS = "SUCCESS"
    FAILED = "FAILED"
    TIMEOUT = "TIMEOUT"
    CANCELLED = "CANCELLED"


class OperationClass(str, Enum):
    READ_ONLY = "READ_ONLY"
    LIGHT_WRITE = "LIGHT_WRITE"
    HEAVY_WRITE = "HEAVY_WRITE"
    DISRUPTIVE = "DISRUPTIVE"


@dataclass(slots=True)
class JobStatus:
    label: str
    elapsed_seconds: float
    completed: bool
    timed_out: bool = False


@dataclass(slots=True)
class JobRecord:
    job_id: str
    host: str
    label: str
    operation_class: OperationClass
    state: JobState = JobState.QUEUED
    elapsed_seconds: float = 0.0
    error: str | None = None


class HostLockRegistry:
    def __init__(self) -> None:
        self._guard = threading.Lock()
        self._locks: dict[str, threading.Lock] = {}

    def _lock_for(self, host: str) -> threading.Lock:
        with self._guard:
            return self._locks.setdefault(host.lower(), threading.Lock())

    @contextmanager
    def hold(self, host: str, operation_class: OperationClass):
        if operation_class == OperationClass.READ_ONLY:
            yield
            return
        lock = self._lock_for(host)
        lock.acquire()
        try:
            yield
        finally:
            lock.release()


class JobManager:
    """Scheduler concorrente com serialização de mutações por estação."""

    def __init__(self, *, max_workers: int = 6, heartbeat_seconds: float = 0.2) -> None:
        self.heartbeat_seconds = max(0.05, float(heartbeat_seconds))
        self._pool = ThreadPoolExecutor(max_workers=max(1, int(max_workers)), thread_name_prefix="central-n2")
        self._locks = HostLockRegistry()
        self._records: dict[str, JobRecord] = {}
        self._guard = threading.Lock()

    def submit(
        self,
        host: str,
        label: str,
        func: Callable[[], T],
        *,
        operation_class: OperationClass = OperationClass.READ_ONLY,
    ) -> tuple[JobRecord, Future[T]]:
        record = JobRecord(uuid.uuid4().hex[:12], host, label, operation_class)
        with self._guard:
            self._records[record.job_id] = record

        def wrapped() -> T:
            started = time.monotonic()
            record.state = JobState.RUNNING
            try:
                with self._locks.hold(host, operation_class):
                    value = func()
                record.state = JobState.SUCCESS
                return value
            except Exception as exc:
                record.state = JobState.FAILED
                record.error = f"{type(exc).__name__}: {exc}"
                raise
            finally:
                record.elapsed_seconds = time.monotonic() - started

        return record, self._pool.submit(wrapped)

    def run_sync(
        self,
        host: str,
        label: str,
        func: Callable[[], T],
        *,
        operation_class: OperationClass = OperationClass.READ_ONLY,
        timeout: float | None = None,
        on_tick: Callable[[JobStatus], None] | None = None,
    ) -> T:
        record, future = self.submit(host, label, func, operation_class=operation_class)
        started = time.monotonic()
        spinner = itertools.cycle("|/-\\")
        while True:
            elapsed = time.monotonic() - started
            remaining = None if timeout is None else timeout - elapsed
            if remaining is not None and remaining <= 0:
                future.cancel()
                record.state = JobState.TIMEOUT
                record.elapsed_seconds = elapsed
                self._clear_status_line()
                if on_tick:
                    on_tick(JobStatus(label, elapsed, False, True))
                raise TimeoutError(f"Operação '{label}' excedeu {timeout:.0f}s")
            wait_for = self.heartbeat_seconds if remaining is None else min(self.heartbeat_seconds, remaining)
            try:
                result = future.result(timeout=wait_for)
                elapsed = time.monotonic() - started
                record.elapsed_seconds = elapsed
                self._clear_status_line()
                if on_tick:
                    on_tick(JobStatus(label, elapsed, True, False))
                return result
            except FutureTimeout:
                elapsed = time.monotonic() - started
                if on_tick:
                    on_tick(JobStatus(label, elapsed, False, False))
                else:
                    sys.stdout.write(f"\r{next(spinner)} {label}... {elapsed:5.1f}s")
                    sys.stdout.flush()

    def list_records(self) -> list[JobRecord]:
        with self._guard:
            return list(self._records.values())

    def shutdown(self) -> None:
        self._pool.shutdown(wait=False, cancel_futures=True)

    @staticmethod
    def _clear_status_line() -> None:
        sys.stdout.write("\r" + " " * 100 + "\r")
        sys.stdout.flush()


class ResponsiveJobRunner:
    """Compatibilidade com a v3 sobre o JobManager v5."""

    def __init__(self, *, heartbeat_seconds: float = 0.2) -> None:
        self._manager = JobManager(max_workers=4, heartbeat_seconds=heartbeat_seconds)

    @property
    def heartbeat_seconds(self) -> float:
        return self._manager.heartbeat_seconds

    def run(
        self,
        label: str,
        func: Callable[[], T],
        *,
        timeout: float | None = None,
        on_tick: Callable[[JobStatus], None] | None = None,
    ) -> T:
        return self._manager.run_sync(
            "local-ui",
            label,
            func,
            operation_class=OperationClass.READ_ONLY,
            timeout=timeout,
            on_tick=on_tick,
        )

    def shutdown(self) -> None:
        self._manager.shutdown()
