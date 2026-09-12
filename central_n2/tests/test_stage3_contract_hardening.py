from __future__ import annotations

import json
from types import SimpleNamespace

from core.capabilities import CapabilityDetector
from core.executor import RemoteExecutor
from core.result import CommandResult
from qualification.models import (
    CaseStatus,
    QualificationCaseResult,
    QualificationReport,
)
from qualification.runner import QualificationRunner
from ui.console_base import ConsoleBase


def _executor_with_result(result: CommandResult) -> RemoteExecutor:
    executor = RemoteExecutor.__new__(RemoteExecutor)

    def execute_powershell(host, script, *, timeout=None, fallback_mode="read_only"):
        return result

    executor.execute_powershell = execute_powershell
    return executor


def test_read_only_json_contract_marks_empty_output_indeterminate():
    result = CommandResult(
        True,
        "powershell",
        "PC01",
        stdout="",
        transport="winrm",
    )
    executor = _executor_with_result(result)

    actual = executor.execute_powershell_json("PC01", "Get-Date")

    assert actual.success is True
    assert actual.indeterminate is True
    assert actual.metadata["structured_output_expected"] is True
    assert actual.metadata["json_parse_error"] is True


def test_mutating_json_contract_marks_invalid_payload_indeterminate():
    result = CommandResult(
        True,
        "powershell",
        "PC01",
        stdout="not-json",
        transport="winrm",
    )
    executor = _executor_with_result(result)

    actual = executor.execute_mutating_powershell_json(
        "PC01",
        "Set-Something",
    )

    assert actual.indeterminate is True
    assert actual.metadata["json_parse_error"] is True


def test_json_contract_accepts_valid_json_null():
    result = CommandResult(
        True,
        "powershell",
        "PC01",
        stdout="null",
        transport="local",
    )
    executor = _executor_with_result(result)

    actual = executor.execute_powershell_json("PC01", "$null")

    assert actual.indeterminate is False
    assert actual.metadata["json_parsed"] is True
    assert actual.data is None


class _CapabilityExecutor:
    def __init__(self, result: CommandResult) -> None:
        self.result = result

    def execute_powershell_json(self, host, script, timeout=None):
        return self.result


def test_capability_probe_rejects_indeterminate_result():
    result = CommandResult(
        True,
        "capabilities",
        "PC01",
        data={},
        transport="winrm",
    ).mark_indeterminate("payload não confirmado")

    report = CapabilityDetector(_CapabilityExecutor(result)).probe("PC01")

    assert report.values == {}
    assert report.error == "payload não confirmado"


def test_capability_probe_rejects_non_dict_payload():
    result = CommandResult(
        True,
        "capabilities",
        "PC01",
        data=["unexpected"],
        transport="winrm",
    )

    report = CapabilityDetector(_CapabilityExecutor(result)).probe("PC01")

    assert report.values == {}
    assert "Payload de capabilities" in str(report.error)


def test_console_does_not_label_indeterminate_as_success(capsys):
    result = CommandResult(
        True,
        "action",
        "PC01",
        transport="winrm",
    ).mark_indeterminate("estado final incerto")

    ConsoleBase.show_result(result)
    output = capsys.readouterr().out

    assert "⚠ INDETERMINADO" in output
    assert "✓ SUCESSO" not in output
    assert "estado final incerto" in output


class _SessionFailureRuntime:
    def open_session(self, host, *, refresh=True):
        raise RuntimeError("endpoint offline")


def test_qualification_exports_evidence_when_session_open_fails(tmp_path):
    runner = QualificationRunner(
        _SessionFailureRuntime(),
        output_dir=tmp_path,
        operator="tester",
    )

    report = runner.run("PC-OFFLINE")

    assert report.passed is False
    assert report.connectivity_state == "ERROR"
    assert len(list(tmp_path.glob("qualification-PC-OFFLINE-*.json"))) == 1
    assert len(list(tmp_path.glob("qualification-PC-OFFLINE-*.md"))) == 1


class _UnknownActionRegistry:
    @staticmethod
    def get(key):
        raise KeyError(key)


class _ReadySession:
    host = "PC01"
    transport = "local"
    ready = True
    connectivity = {
        "state": "READY_LOCAL",
        "diagnosis": "local",
    }
    capabilities = {
        "DomainMember": False,
        "PrinterManagement": False,
        "GLPI": False,
        "WindowsUpdateCOM": False,
        "Battery": False,
    }
    capability_error = None


class _UnknownActionRuntime:
    def __init__(self) -> None:
        self.session = _ReadySession()
        self.registry = _UnknownActionRegistry()
        self.sessions = SimpleNamespace(get=lambda host: self.session)

    def open_session(self, host, *, refresh=True):
        return self.session

    @staticmethod
    def probe_handlers():
        def ok(host):
            return CommandResult(
                True,
                "probe",
                host,
                stdout='{"ok":true}',
                data={"ok": True},
                transport="local",
            )

        return {
            "core.health": ok,
            "inventory.processes": ok,
            "inventory.services": ok,
            "network.adapters": ok,
            "network.ip": ok,
            "devices.problems": ok,
            "security.posture": ok,
        }


def test_unknown_qualification_action_becomes_evidenced_failure(tmp_path):
    runner = QualificationRunner(
        _UnknownActionRuntime(),
        output_dir=tmp_path,
    )

    report = runner.run(
        "PC01",
        requested_actions=["missing.action"],
    )

    action = next(
        case
        for case in report.cases
        if case.key == "action.missing.action"
    )
    assert action.status is CaseStatus.FAIL
    assert report.passed is False
    assert "não encontrada" in action.message
    assert len(list(tmp_path.glob("qualification-PC01-*.json"))) == 1


def test_qualification_export_redacts_secret_material(tmp_path):
    secret = "do-not-persist-me"
    report = QualificationReport(
        host="PC01",
        profile="auto",
        started_at="2026-09-12T00:00:00+00:00",
        finished_at="2026-09-12T00:00:01+00:00",
        transport="local",
        connectivity_state="READY_LOCAL",
        cases=[
            QualificationCaseResult(
                key="core.test",
                title="Teste",
                category="Core",
                status=CaseStatus.FAIL,
                started_at="2026-09-12T00:00:00+00:00",
                finished_at="2026-09-12T00:00:01+00:00",
                duration_ms=1,
                message=f"token={secret}",
                evidence={
                    "password": secret,
                    "nested": {"authorization": f"Bearer {secret}"},
                },
            )
        ],
    )
    runner = QualificationRunner(
        SimpleNamespace(),
        output_dir=tmp_path,
    )

    runner.export(report)

    json_text = next(tmp_path.glob("*.json")).read_text(encoding="utf-8")
    md_text = next(tmp_path.glob("*.md")).read_text(encoding="utf-8")
    payload = json.loads(json_text)

    assert secret not in json_text
    assert secret not in md_text
    assert payload["cases"][0]["evidence"]["password"] == "***"
