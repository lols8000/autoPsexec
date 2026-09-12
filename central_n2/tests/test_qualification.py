from __future__ import annotations

import json
from types import SimpleNamespace

from core.result import CommandResult
from execution import DisconnectMode
from qualification.cli import parse_action_parameters
from qualification.models import CaseStatus
from qualification.runner import QualificationRunner


class FakeSession:
    def __init__(
        self,
        *,
        ready: bool = True,
        transport: str = "winrm",
        capabilities: dict | None = None,
    ) -> None:
        self.host = "PC01"
        self.transport = transport
        self.connectivity = {
            "state": "READY_WINRM" if ready else "NO_USABLE_TRANSPORT",
            "diagnosis": "ok" if ready else "sem transporte",
        }
        self.capabilities = (
            capabilities
            if capabilities is not None
            else {
                "DomainMember": True,
                "PrinterManagement": True,
                "GLPI": True,
                "WindowsUpdateCOM": True,
                "Battery": True,
            }
        )
        self.capability_error = None
        self._ready = ready

    @property
    def ready(self) -> bool:
        return self._ready


class FakeRegistry:
    def __init__(self, *, disruptive: bool = False) -> None:
        self.disruptive = disruptive

    def get(self, key: str):
        spec = SimpleNamespace(
            key=key,
            title="Teste",
            category_label="Teste",
            may_break_connectivity=self.disruptive,
            disconnect_mode=(
                DisconnectMode.TEMPORARY
                if self.disruptive
                else DisconnectMode.NONE
            ),
            risk=SimpleNamespace(value="HIGH" if self.disruptive else "LOW"),
        )
        return SimpleNamespace(spec=spec)


class FakeRuntime:
    def __init__(self, session: FakeSession | None = None) -> None:
        self.session = session or FakeSession()
        self.sessions = SimpleNamespace(get=lambda host: self.session)
        self.registry = FakeRegistry()
        self.engine = SimpleNamespace(execute=self._unexpected_execute)

    @staticmethod
    def _unexpected_execute(*args, **kwargs):
        raise AssertionError("engine não deveria ser executado neste teste")

    def open_session(self, host: str, *, refresh: bool = True):
        return self.session

    @staticmethod
    def policy_context(session):
        return SimpleNamespace()

    @staticmethod
    def recover(host: str, timeout_seconds: int, delay_seconds: int):
        return None

    def probe_handlers(self):
        def ok(host: str):
            return CommandResult(
                True,
                "probe",
                host,
                data={"ok": True},
                transport=self.session.transport,
            )

        return {
            "core.health": ok,
            "inventory.processes": ok,
            "inventory.services": ok,
            "network.adapters": ok,
            "network.ip": ok,
            "devices.problems": ok,
            "security.posture": ok,
            "domain.status": ok,
            "printers.inventory": ok,
            "glpi.status": ok,
            "updates.status": ok,
        }


def test_parse_action_parameters_groups_values_by_action():
    result = parse_action_parameters(
        [
            "network.adapter_restart:adapter_name=Ethernet",
            "energy.restart:delay_seconds=5",
        ]
    )

    assert result == {
        "network.adapter_restart": {"adapter_name": "Ethernet"},
        "energy.restart": {"delay_seconds": "5"},
    }


def test_parse_action_parameters_rejects_invalid_syntax():
    try:
        parse_action_parameters(["broken"])
    except Exception as exc:
        assert "ACTION:KEY=VALUE" in str(exc)
    else:
        raise AssertionError("Sintaxe inválida deveria falhar")


def test_baseline_matrix_exports_json_and_markdown(tmp_path):
    runner = QualificationRunner(
        FakeRuntime(),
        output_dir=tmp_path,
        operator="tester",
    )

    report = runner.run("PC01", profile="winrm")

    assert report.passed is True
    assert report.summary()["FAIL"] == 0
    assert any(case.key == "core.health" for case in report.cases)
    assert len(list(tmp_path.glob("qualification-PC01-*.json"))) == 1
    assert len(list(tmp_path.glob("qualification-PC01-*.md"))) == 1

    payload = json.loads(
        list(tmp_path.glob("qualification-PC01-*.json"))[0].read_text(
            encoding="utf-8"
        )
    )
    assert payload["host"] == "PC01"
    assert payload["transport"] == "winrm"


def test_profile_transport_mismatch_is_reported(tmp_path):
    runner = QualificationRunner(
        FakeRuntime(FakeSession(transport="psexec")),
        output_dir=tmp_path,
    )

    report = runner.run("PC01", profile="winrm")
    profile_case = next(
        case
        for case in report.cases
        if case.key == "core.profile_expectation"
    )

    assert profile_case.status is CaseStatus.FAIL
    assert "winrm" in profile_case.message
    assert "psexec" in profile_case.message
    assert report.passed is False


def test_optional_capability_cases_are_skipped(tmp_path):
    session = FakeSession(
        capabilities={
            "DomainMember": False,
            "PrinterManagement": False,
            "GLPI": False,
            "WindowsUpdateCOM": False,
            "Battery": False,
        }
    )
    runner = QualificationRunner(
        FakeRuntime(session),
        output_dir=tmp_path,
    )

    report = runner.run("PC01")

    skipped = {
        case.key
        for case in report.cases
        if case.status is CaseStatus.SKIP
    }
    assert "domain.status" in skipped
    assert "printers.inventory" in skipped
    assert "glpi.status" in skipped
    assert "updates.status" in skipped
    assert report.passed is True


def test_specialized_printer_profile_requires_printer_capability(tmp_path):
    session = FakeSession(
        capabilities={
            "DomainMember": False,
            "PrinterManagement": False,
            "GLPI": False,
            "WindowsUpdateCOM": False,
            "Battery": False,
        }
    )
    runner = QualificationRunner(FakeRuntime(session), output_dir=tmp_path)

    report = runner.run("PC01", profile="printer")

    assert report.passed is False
    printer_case = next(
        case for case in report.cases if case.key == "printers.inventory"
    )
    assert printer_case.status is CaseStatus.SKIP


def test_specialized_glpi_profile_requires_glpi_capability(tmp_path):
    session = FakeSession(
        capabilities={
            "DomainMember": True,
            "PrinterManagement": True,
            "GLPI": False,
            "WindowsUpdateCOM": True,
            "Battery": False,
        }
    )
    runner = QualificationRunner(FakeRuntime(session), output_dir=tmp_path)

    report = runner.run("PC01", profile="glpi")

    assert report.passed is False


def test_specialized_domain_profile_requires_domain_membership(tmp_path):
    session = FakeSession(
        capabilities={
            "DomainMember": False,
            "PrinterManagement": True,
            "GLPI": True,
            "WindowsUpdateCOM": True,
            "Battery": False,
        }
    )
    runner = QualificationRunner(FakeRuntime(session), output_dir=tmp_path)

    report = runner.run("PC01", profile="domain")

    assert report.passed is False


def test_specialized_notebook_profile_requires_battery(tmp_path):
    session = FakeSession(
        capabilities={
            "DomainMember": True,
            "PrinterManagement": True,
            "GLPI": True,
            "WindowsUpdateCOM": True,
            "Battery": False,
        }
    )
    runner = QualificationRunner(FakeRuntime(session), output_dir=tmp_path)

    report = runner.run("PC01", profile="notebook")

    assert report.passed is False


def test_specialized_notebook_profile_passes_with_battery(tmp_path):
    runner = QualificationRunner(FakeRuntime(), output_dir=tmp_path)

    report = runner.run("PC01", profile="notebook")

    assert report.passed is True


def test_no_transport_fails_connectivity_and_skips_probes(tmp_path):
    runner = QualificationRunner(
        FakeRuntime(FakeSession(ready=False, capabilities={})),
        output_dir=tmp_path,
    )

    report = runner.run("PC01")

    connectivity = next(
        case for case in report.cases if case.key == "core.connectivity"
    )
    health = next(
        case for case in report.cases if case.key == "core.health"
    )
    assert connectivity.status is CaseStatus.FAIL
    assert health.status is CaseStatus.SKIP
    assert report.passed is False


def test_disruptive_action_requires_explicit_authorization(tmp_path):
    runtime = FakeRuntime()
    runtime.registry = FakeRegistry(disruptive=True)
    runner = QualificationRunner(runtime, output_dir=tmp_path)

    report = runner.run(
        "PC01",
        requested_actions=["network.adapter_restart"],
        action_parameters={
            "network.adapter_restart": {"adapter_name": "Ethernet"}
        },
        allow_disruptive=False,
    )

    action = next(
        case
        for case in report.cases
        if case.key == "action.network.adapter_restart"
    )
    assert action.status is CaseStatus.FAIL
    assert "não autorizada" in action.message


def test_disruptive_action_requires_matching_host_confirmation(tmp_path):
    runtime = FakeRuntime()
    runtime.registry = FakeRegistry(disruptive=True)
    runner = QualificationRunner(runtime, output_dir=tmp_path)

    report = runner.run(
        "PC01",
        requested_actions=["energy.restart"],
        allow_disruptive=True,
        confirmed_host="OTHER-PC",
    )

    action = next(
        case
        for case in report.cases
        if case.key == "action.energy.restart"
    )
    assert action.status is CaseStatus.FAIL
    assert "Confirmação nominal" in action.message
