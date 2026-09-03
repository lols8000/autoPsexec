from __future__ import annotations

import time

import pytest

from core.jobs import JobStatus, ResponsiveJobRunner
from modules.disk import DiskModule
from modules.domain import DomainModule
from modules.printers import PrintersModule
from modules.security import SecurityModule


class DummyExecutor:
    def execute_cmd(self, *args, **kwargs):
        return (args, kwargs)

    def execute_powershell_json(self, *args, **kwargs):
        return (args, kwargs)

    def execute_remote_powershell_with_fallback(self, *args, **kwargs):
        return (args, kwargs)


def test_responsive_runner_returns_value_and_ticks():
    runner = ResponsiveJobRunner(heartbeat_seconds=0.05)
    ticks: list[JobStatus] = []
    try:
        value = runner.run("teste", lambda: (time.sleep(0.12), 42)[1], timeout=1, on_tick=ticks.append)
    finally:
        runner.shutdown()
    assert value == 42
    assert any(not tick.completed for tick in ticks)
    assert ticks[-1].completed is True


def test_responsive_runner_timeout():
    runner = ResponsiveJobRunner(heartbeat_seconds=0.02)
    try:
        with pytest.raises(TimeoutError):
            runner.run("lento", lambda: time.sleep(0.2), timeout=0.05, on_tick=lambda _: None)
    finally:
        runner.shutdown()


def test_v3_module_aliases_are_available():
    executor = DummyExecutor()
    assert callable(PrintersModule(executor).list_printers)
    assert callable(DomainModule(executor).gpupdate)
    assert callable(DiskModule(executor).space)
    assert callable(DiskModule(executor).profile_sizes)
    assert callable(SecurityModule(executor).posture)
