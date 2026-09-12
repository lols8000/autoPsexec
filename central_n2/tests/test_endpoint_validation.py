from __future__ import annotations

import argparse
import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from core.result import CommandResult
from core.session import WorkstationSession
from execution import RiskLevel
from validation.cli import ACKNOWLEDGEMENT, _validate_flags
from validation.models import (
    CampaignResult,
    CheckState,
    ControlledActionSpec,
    EndpointSpec,
    EndpointValidationResult,
    ValidationCheck,
    fingerprint_target,
)
from validation.runner import (
    EndpointValidationRunner,
    _check_from_result,
    load_campaign,
    run_campaign,
    write_reports,
)


def _session(
    *,
    state: str = "READY_WINRM",
    transport: str = "winrm",
    capabilities: dict | None = None,
) -> WorkstationSession:
    values = {
        "IsAdmin": True,
        "DomainMember": True,
        "Battery": True,
    }
    if capabilities:
        values.update(capabilities)
    return WorkstationSession(
        host="PRIVATE-HOST",
        transport=transport,
        opened_at="2026-09-12T00:00:00+00:00",
        connectivity={
            "state": state,
            "dns": True,
            "tcp_445": True,
            "tcp_5985": transport == "winrm",
            "tcp_5986": False,
        },
        capabilities=values,
    )


class _ReadModule:
    def __getattr__(self, name):
        return lambda *args, **kwargs: CommandResult(
            True,
            name,
            "PRIVATE-HOST",
            transport="winrm",
            data={"ok": True},
        )


class _Sessions:
    def __init__(self, session):
        self.session = session

    def open(self, host, refresh=False):
        return self.session


def _runner_for_safe_checks(session: WorkstationSession):
    runner = EndpointValidationRunner.__new__(
        EndpointValidationRunner
    )
    runner.health = _ReadModule()
    runner.network = _ReadModule()
    runner.printers = _ReadModule()
    runner.domain = _ReadModule()
    runner.glpi = _ReadModule()
    runner.devices = _ReadModule()
    runner.sessions = _Sessions(session)
    return runner


def test_fingerprint_is_stable_and_does_not_expose_target():
    target = "INTERNAL-WORKSTATION-123"

    value = fingerprint_target(target)

    assert value == fingerprint_target(target.lower())
    assert len(value) == 12
    assert target not in value


def test_endpoint_status_prefers_fail_then_unknown():
    base = dict(
        alias="A",
        target_fingerprint="abc",
        correlation_id="corr",
        started_at="start",
        finished_at="end",
    )
    passed = EndpointValidationResult(
        **base,
        checks=[
            ValidationCheck("a", CheckState.PASS, "ok"),
            ValidationCheck(
                "skip",
                CheckState.SKIP,
                "skip",
                required=False,
            ),
        ],
    )
    unknown = EndpointValidationResult(
        **base,
        checks=[
            ValidationCheck("a", CheckState.UNKNOWN, "unknown"),
        ],
    )
    failed = EndpointValidationResult(
        **base,
        checks=[
            ValidationCheck("a", CheckState.UNKNOWN, "unknown"),
            ValidationCheck("b", CheckState.FAIL, "fail"),
        ],
    )

    assert passed.status is CheckState.PASS
    assert unknown.status is CheckState.UNKNOWN
    assert failed.status is CheckState.FAIL


def test_public_payload_never_contains_target():
    target = "TOP-SECRET-HOSTNAME"
    endpoint = EndpointValidationResult(
        alias="WINRM-LAB",
        target_fingerprint=fingerprint_target(target),
        correlation_id="abc",
        started_at="start",
        finished_at="end",
        checks=[
            ValidationCheck(
                "session.ready",
                CheckState.PASS,
                "ok",
            )
        ],
    )

    encoded = json.dumps(endpoint.public_dict())

    assert target not in encoded
    assert "WINRM-LAB" in encoded


def test_load_campaign_parses_actions_and_disabled_endpoints(
    tmp_path: Path,
):
    path = tmp_path / "matrix.json"
    path.write_text(
        json.dumps(
            {
                "name": "field",
                "endpoints": [
                    {
                        "alias": "A",
                        "target": "host-a",
                        "enabled": False,
                        "roles": ["domain"],
                        "actions": [
                            {
                                "key": "network.flush_dns",
                                "parameters": {},
                                "rollback_after": False,
                            }
                        ],
                    }
                ],
            }
        ),
        encoding="utf-8",
    )

    name, endpoints = load_campaign(path)

    assert name == "field"
    assert len(endpoints) == 1
    assert endpoints[0].enabled is False
    assert endpoints[0].roles == ("domain",)
    assert endpoints[0].actions[0].key == "network.flush_dns"


def test_load_campaign_rejects_duplicate_alias(tmp_path: Path):
    path = tmp_path / "matrix.json"
    path.write_text(
        json.dumps(
            {
                "endpoints": [
                    {"alias": "A", "target": "one"},
                    {"alias": "a", "target": "two"},
                ]
            }
        ),
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Alias duplicado"):
        load_campaign(path)


def test_check_from_result_maps_indeterminate_to_unknown():
    failed = CommandResult.failure(
        "PC",
        "probe",
        "timeout",
    ).mark_indeterminate("uncertain")

    check = _check_from_result("probe", failed, "ok")

    assert check.state is CheckState.UNKNOWN


def test_safe_checks_validate_transport_capabilities_and_roles():
    session = _session(
        capabilities={
            "PrinterManagement": True,
            "GLPI": True,
            "PnpDevice": True,
        }
    )
    runner = _runner_for_safe_checks(session)
    spec = EndpointSpec(
        alias="LAB",
        target="PRIVATE-HOST",
        expected_state="READY_WINRM",
        expected_transport="winrm",
        required_capabilities=(
            "IsAdmin",
            "PrinterManagement",
        ),
        roles=("domain", "printer", "glpi", "pnp", "laptop"),
    )

    checks = runner._safe_checks(spec, session)

    assert checks
    assert all(item.state is CheckState.PASS for item in checks)
    keys = {item.key for item in checks}
    assert "probe.health" in keys
    assert "probe.domain" in keys
    assert "probe.printers" in keys
    assert "probe.glpi" in keys
    assert "probe.pnp" in keys


def test_safe_checks_stop_after_unusable_transport():
    session = _session(
        state="NO_USABLE_TRANSPORT",
        transport="unknown",
    )
    runner = _runner_for_safe_checks(session)
    spec = EndpointSpec(
        alias="BROKEN",
        target="PRIVATE-HOST",
    )

    checks = runner._safe_checks(spec, session)

    assert len(checks) == 1
    assert checks[0].key == "session.ready"
    assert checks[0].state is CheckState.FAIL


def test_controlled_actions_are_skipped_without_explicit_permission():
    runner = EndpointValidationRunner.__new__(
        EndpointValidationRunner
    )
    spec = EndpointSpec(
        alias="LAB",
        target="PRIVATE-HOST",
        actions=(
            ControlledActionSpec("network.flush_dns"),
        ),
    )

    checks = runner._controlled_checks(
        spec,
        _session(),
        allow_mutations=False,
        allow_high_risk=False,
        allow_reboot=False,
        operator="tester",
    )

    assert len(checks) == 1
    assert checks[0].state is CheckState.SKIP
    assert checks[0].required is False


def test_high_risk_action_is_blocked_without_flag():
    runner = EndpointValidationRunner.__new__(
        EndpointValidationRunner
    )
    runner.registry = SimpleNamespace(
        get=lambda key: SimpleNamespace(
            spec=SimpleNamespace(
                risk=RiskLevel.HIGH,
                destructive=False,
            )
        )
    )
    spec = EndpointSpec(
        alias="LAB",
        target="PRIVATE-HOST",
        actions=(ControlledActionSpec("test.high"),),
    )

    checks = runner._controlled_checks(
        spec,
        _session(),
        allow_mutations=True,
        allow_high_risk=False,
        allow_reboot=False,
        operator="tester",
    )

    assert checks[0].state is CheckState.FAIL
    assert "--allow-high-risk" in checks[0].message


def test_destructive_action_is_never_automated():
    runner = EndpointValidationRunner.__new__(
        EndpointValidationRunner
    )
    runner.registry = SimpleNamespace(
        get=lambda key: SimpleNamespace(
            spec=SimpleNamespace(
                risk=RiskLevel.MEDIUM,
                destructive=True,
            )
        )
    )
    spec = EndpointSpec(
        alias="LAB",
        target="PRIVATE-HOST",
        actions=(ControlledActionSpec("file.remove"),),
    )

    checks = runner._controlled_checks(
        spec,
        _session(),
        allow_mutations=True,
        allow_high_risk=True,
        allow_reboot=False,
        operator="tester",
    )

    assert checks[0].state is CheckState.FAIL
    assert "destrutiva" in checks[0].message


def test_mutation_flags_require_exact_acknowledgement():
    args = argparse.Namespace(
        allow_mutations=True,
        allow_high_risk=False,
        allow_reboot=False,
        acknowledge="wrong",
    )

    with pytest.raises(ValueError, match="confirmação explícita"):
        _validate_flags(args)

    args.acknowledge = ACKNOWLEDGEMENT
    _validate_flags(args)


def test_run_campaign_ignores_disabled_endpoint():
    calls = []

    class Runner:
        def validate_endpoint(self, endpoint, **kwargs):
            calls.append(endpoint.alias)
            return EndpointValidationResult(
                alias=endpoint.alias,
                target_fingerprint="abc",
                correlation_id="corr",
                started_at="start",
                finished_at="end",
                checks=[
                    ValidationCheck(
                        "ok",
                        CheckState.PASS,
                        "ok",
                    )
                ],
            )

    campaign = run_campaign(
        Runner(),
        "test",
        [
            EndpointSpec("A", "host-a", enabled=True),
            EndpointSpec("B", "host-b", enabled=False),
        ],
    )

    assert calls == ["A"]
    assert campaign.status is CheckState.PASS


def test_write_reports_are_sanitized(tmp_path: Path):
    private_target = "REAL-CORPORATE-HOST"
    endpoint = EndpointValidationResult(
        alias="LAB",
        target_fingerprint=fingerprint_target(private_target),
        correlation_id="corr",
        started_at="start",
        finished_at="end",
        checks=[
            ValidationCheck(
                "session.ready",
                CheckState.PASS,
                "ready",
            )
        ],
    )
    campaign = CampaignResult(
        name="field",
        started_at="start",
        finished_at="end",
        endpoints=[endpoint],
    )

    json_path, md_path = write_reports(campaign, tmp_path)

    combined = (
        json_path.read_text(encoding="utf-8")
        + md_path.read_text(encoding="utf-8")
    )
    assert private_target not in combined
    assert "LAB" in combined
