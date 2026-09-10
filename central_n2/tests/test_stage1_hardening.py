from __future__ import annotations

from pathlib import Path

from core.context import AttendanceContext
from core.executor import RemoteExecutor
from core.result import CommandResult
from core.transport.winrm import WinRMTransport
from core.validation import validate_port, validate_sid
from modules.diagnostic_package import DiagnosticPackageModule
from modules.updates import UpdatesModule


def test_attendance_context_never_leaks_previous_host_state(tmp_path: Path):
    context = AttendanceContext.start("PC-A")
    context.health_snapshot = {"Hostname": "PC-A"}
    context.diagnoses.append("diagnosis-a")
    context.playbook = "playbook-a"
    context.remediation = "remediation-a"
    context.report_path = tmp_path / "pc-a.md"

    context.bind("PC-B", session="session-b")

    assert context.belongs_to("PC-B")
    assert not context.belongs_to("PC-A")
    assert context.health_snapshot is None
    assert context.diagnoses == []
    assert context.playbook is None
    assert context.remediation is None
    assert context.report_path is None


def test_winrm_cmd_uses_remote_exit_code():
    def runner(cmd, *, host, action, timeout=None, output_encoding=None):
        return CommandResult(
            True,
            "wrapper",
            host,
            stdout='{"ExitCode":5,"Output":"remote failure"}',
            transport="local",
            duration_ms=12,
        )

    transport = WinRMTransport(runner, lambda: "")
    result = transport.execute_cmd("PC01", "tool.exe /arg")

    assert result.success is False
    assert result.return_code == 5
    assert result.transport == "winrm"
    assert result.metadata["remote_exit_code_confirmed"] is True


def test_winrm_cmd_unparseable_result_is_indeterminate():
    def runner(cmd, *, host, action, timeout=None, output_encoding=None):
        return CommandResult(
            True,
            "wrapper",
            host,
            stdout="not-json",
            transport="local",
        )

    result = WinRMTransport(runner, lambda: "").execute_cmd("PC01", "hostname")

    assert result.success is False
    assert result.return_code == 125
    assert result.indeterminate is True


class _AvailablePsExec:
    def available(self):
        return True


class _TransportManager:
    def __init__(self):
        self.invalidated = []

    def invalidate(self, host):
        self.invalidated.append(host)


def _executor_for_fallback():
    executor = RemoteExecutor.__new__(RemoteExecutor)
    executor.psexec_transport = _AvailablePsExec()
    executor.transport_manager = _TransportManager()
    return executor


def test_mutation_does_not_fallback_after_indeterminate_transport_loss():
    executor = _executor_for_fallback()
    original = CommandResult.failure(
        "PC01",
        "restart-service",
        "WinRMOperationTimeout PSSessionStateBroken",
        return_code=124,
        transport="winrm",
    )
    called = []

    result = executor._fallback_after_winrm_failure(
        host="PC01",
        result=original,
        mode="mutation",
        fallback=lambda: called.append(True),
    )

    assert result is original
    assert result.indeterminate is True
    assert called == []
    assert result.metadata["fallback_suppressed"] is True


def test_mutation_can_fallback_when_failure_is_definitely_pre_execution():
    executor = _executor_for_fallback()
    original = CommandResult.failure(
        "10.0.0.10",
        "action",
        "CannotUseIPAddress: destino exige TrustedHosts",
        transport="winrm",
    )
    fallback_result = CommandResult(
        True,
        "action",
        "10.0.0.10",
        transport="psexec",
    )

    result = executor._fallback_after_winrm_failure(
        host="10.0.0.10",
        result=original,
        mode="mutation",
        fallback=lambda: fallback_result,
    )

    assert result is fallback_result
    assert result.metadata["fallback_from"] == "winrm"
    assert result.metadata["fallback_reason"] == "pre_execution"


class _DiagnosticExecutor:
    def execute_powershell_json(self, host, script, timeout=None):
        return CommandResult(
            True,
            "diagnostic",
            host,
            data={"Computer": {"Hostname": host}},
            transport="local",
        )


def test_diagnostic_package_preserves_command_result_contract(tmp_path: Path):
    module = DiagnosticPackageModule(_DiagnosticExecutor(), tmp_path)
    result = module.collect("PC01")

    assert isinstance(result, CommandResult)
    assert result.success is True
    assert Path(result.data["path"]).exists()
    assert result.metadata["artifact_path"] == result.data["path"]


class _UpdateExecutor:
    def __init__(self):
        self.script = ""

    def execute_mutating_powershell_json(self, host, script, timeout=None):
        self.script = script
        return CommandResult(True, "update-reset", host, transport="local")


def test_windows_update_reset_has_finally_service_restoration():
    executor = _UpdateExecutor()
    result = UpdatesModule(executor).reset_components("PC01")

    assert result.success is True
    assert "finally" in executor.script
    assert "Start-Service" in executor.script
    assert "$original" in executor.script
    assert result.metadata["transactional_cleanup"] is True


def test_validation_rejects_invalid_port_and_sid():
    assert validate_port(445) == 445
    assert validate_sid("S-1-5-21-1-2-3-1001")
    try:
        validate_port(70000)
    except ValueError:
        pass
    else:
        raise AssertionError("porta inválida deveria falhar")

    try:
        validate_sid("S-1-INVALID")
    except ValueError:
        pass
    else:
        raise AssertionError("SID inválido deveria falhar")


def test_winrm_probe_validates_invoke_command():
    captured = {}

    def runner(cmd, *, host, action, timeout=None, output_encoding=None):
        captured["command"] = " ".join(cmd)
        captured["action"] = action
        return CommandResult(
            True,
            "probe",
            host,
            stdout="CENTRAL_N2_WINRM_OK",
        )

    result = WinRMTransport(runner, lambda: "").test("PC01")

    assert result.success is True
    assert captured["action"] == "test_winrm"
    assert "Invoke-Command" in captured["command"]
    assert "Test-WSMan" in captured["command"]
