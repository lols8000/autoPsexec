from __future__ import annotations

import socket
import threading
import time

import pytest

from core.config import deep_merge
from core.connectivity import ConnectivityDiagnostics
from core.host_identity import HostIdentity
from core.jobs import JobManager, JobState, OperationClass
from core.result import CommandResult
from core.transport.manager import TransportManager
from storage.database import CentralDatabase


class FakeTransport:
    def __init__(self, name: str, success: bool = True, available: bool = True):
        self.name = name
        self._success = success
        self._available = available
        self.test_calls = 0

    def available(self):
        return self._available

    def test(self, host):
        self.test_calls += 1
        return CommandResult(self._success, "test", host, transport=self.name)

    def execute_powershell(self, host, script, timeout=None):
        return CommandResult(True, script, host, transport=self.name)

    def execute_cmd(self, host, command, timeout=None):
        return CommandResult(True, command, host, transport=self.name)


def test_deep_merge_keeps_base_and_overrides_nested():
    base = {"a": 1, "nested": {"x": 1, "y": 2}}
    override = {"nested": {"x": 9}, "b": 3}
    assert deep_merge(base, override) == {"a": 1, "nested": {"x": 9, "y": 2}, "b": 3}


def test_host_identity_detects_localhost_and_hostname():
    assert HostIdentity.is_local("localhost")
    assert HostIdentity.is_local("127.0.0.1")
    assert HostIdentity.is_local(socket.gethostname())


def test_transport_manager_uses_local_without_testing_winrm():
    local = FakeTransport("local")
    winrm = FakeTransport("winrm")
    psexec = FakeTransport("psexec")
    manager = TransportManager(local, winrm, psexec)
    assert manager.select("localhost").name == "local"
    assert winrm.test_calls == 0


def test_transport_manager_prefers_winrm_then_psexec():
    local = FakeTransport("local")
    winrm = FakeTransport("winrm", success=True)
    psexec = FakeTransport("psexec")
    manager = TransportManager(local, winrm, psexec)
    assert manager.select("remote.invalid").name == "winrm"

    winrm2 = FakeTransport("winrm", success=False)
    manager2 = TransportManager(local, winrm2, psexec)
    assert manager2.select("remote2.invalid").name == "psexec"


def test_job_manager_serializes_writes_per_host():
    manager = JobManager(max_workers=2, heartbeat_seconds=0.01)
    order = []
    active = 0
    max_active = 0
    guard = threading.Lock()

    def work(tag):
        nonlocal active, max_active
        with guard:
            active += 1
            max_active = max(max_active, active)
            order.append(("start", tag))
        time.sleep(0.05)
        with guard:
            order.append(("end", tag))
            active -= 1
        return tag

    try:
        _, f1 = manager.submit("PC01", "a", lambda: work("a"), operation_class=OperationClass.HEAVY_WRITE)
        _, f2 = manager.submit("PC01", "b", lambda: work("b"), operation_class=OperationClass.HEAVY_WRITE)
        assert f1.result(timeout=1) == "a"
        assert f2.result(timeout=1) == "b"
        assert max_active == 1
    finally:
        manager.shutdown()


class LocalOnlyExecutor:
    psexec_path = None

    def test_winrm(self, host):
        raise AssertionError("WinRM não deve ser consultado para alvo local")

    def test_admin_share(self, host):
        raise AssertionError("ADMIN$ não deve ser consultado para alvo local")

    def ping(self, host):
        raise AssertionError("Ping externo não deve ser necessário para alvo local")

    def select_transport(self, host, refresh=False):
        return "local"


def test_connectivity_local_skips_remote_probes():
    report = ConnectivityDiagnostics(LocalOnlyExecutor()).run("localhost")
    assert report["is_local"] is True
    assert report["selected_transport"] == "local"
    assert report["winrm"] is None
    assert report["admin_share"] is None


def test_job_observer_persists_final_state(tmp_path):
    database = CentralDatabase(tmp_path / "central.db")
    manager = JobManager(
        max_workers=1,
        heartbeat_seconds=0.01,
        observer=database.save_job,
    )
    try:
        record, future = manager.submit(
            "PC01",
            "coleta",
            lambda: "ok",
            operation_class=OperationClass.READ_ONLY,
        )
        assert future.result(timeout=1) == "ok"
        rows = database.recent_jobs("PC01", limit=1)
        assert rows[0]["job_id"] == record.job_id
        assert rows[0]["state"] == JobState.SUCCESS.value
    finally:
        manager.shutdown()


def test_timeout_state_is_not_overwritten_by_late_worker():
    manager = JobManager(max_workers=1, heartbeat_seconds=0.01)
    try:
        with pytest.raises(TimeoutError):
            manager.run_sync(
                "PC01",
                "demorado",
                lambda: time.sleep(0.08),
                timeout=0.02,
            )
        time.sleep(0.10)
        record = manager.list_records()[-1]
        assert record.state == JobState.TIMEOUT
    finally:
        manager.shutdown()
