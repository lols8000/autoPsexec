from __future__ import annotations

import json
from pathlib import Path

from core.baselines import BaselineRepository
from core.evaluation import EvaluationState
from diagnostics.engine import DiagnosticEngine
from modules.compliance import evaluate_compliance
from playbooks import PlaybookAnalyzer
from playbooks.base import PlaybookExecution


def test_missing_metrics_are_unknown_not_false_failures():
    report = evaluate_compliance(
        {
            "DiskFreePercent": None,
            "UptimeDays": None,
            "PendingReboot": None,
            "StoppedAutoServices": None,
            "DefenderEnabled": None,
            "FirewallEnabled": None,
            "GlpiRunning": None,
        },
        {
            "min_disk_free_percent": 15,
            "max_uptime_days": 30,
            "defender_required": True,
            "firewall_required": True,
            "glpi_required": True,
            "pending_reboot_not_allowed": True,
        },
    )

    assert report["failed"] == 0
    assert report["unknown"] > 0
    assert report["overall_state"] == EvaluationState.UNKNOWN.value
    assert all(
        item["state"] != EvaluationState.FAIL.value
        for item in report["items"]
    )


def test_selected_baseline_is_not_overridden_by_default_config(tmp_path: Path):
    (tmp_path / "DEFAULT.json").write_text(
        json.dumps(
            {
                "profile": "DEFAULT",
                "min_disk_free_percent": 15,
                "max_uptime_days": 30,
            }
        ),
        encoding="utf-8",
    )
    (tmp_path / "TI.json").write_text(
        json.dumps(
            {
                "profile": "TI",
                "min_disk_free_percent": 20,
                "max_uptime_days": 14,
            }
        ),
        encoding="utf-8",
    )
    repo = BaselineRepository(tmp_path)

    effective = repo.resolve(
        "TI",
        {"profile": "DEFAULT", "overrides": {}},
    )

    assert effective["min_disk_free_percent"] == 20
    assert effective["max_uptime_days"] == 14


def test_explicit_baseline_override_wins(tmp_path: Path):
    (tmp_path / "DEFAULT.json").write_text(
        json.dumps(
            {
                "profile": "DEFAULT",
                "min_disk_free_percent": 15,
                "max_uptime_days": 30,
            }
        ),
        encoding="utf-8",
    )
    (tmp_path / "TI.json").write_text(
        json.dumps(
            {
                "profile": "TI",
                "min_disk_free_percent": 20,
                "max_uptime_days": 14,
            }
        ),
        encoding="utf-8",
    )
    repo = BaselineRepository(tmp_path)

    effective = repo.resolve(
        "TI",
        {
            "profile": "TI",
            "overrides": {
                "max_uptime_days": 7,
            },
        },
    )

    assert effective["min_disk_free_percent"] == 20
    assert effective["max_uptime_days"] == 7


def test_diagnostic_engine_accepts_cpuaverage_key():
    findings = DiagnosticEngine().evaluate(
        {
            "CPUAverage": 95,
            "RAMUsedPercent": 50,
            "DiskFreePercent": 50,
            "UptimeDays": 1,
            "PendingReboot": False,
            "StoppedAutoServices": 0,
            "DefenderEnabled": True,
            "FirewallEnabled": True,
            "GlpiRunning": True,
        },
        {"min_disk_free_percent": 15},
    )
    assert any(item.id == "CPU_PRESSURE" for item in findings)


def test_network_playbook_produces_actionable_findings():
    execution = PlaybookExecution(
        playbook="network",
        host="PC01",
        started_at="start",
        finished_at="finish",
        steps=[
            {
                "key": "network",
                "label": "IP",
                "success": True,
                "data": [
                    {
                        "InterfaceAlias": "Ethernet",
                        "IPv4": "",
                        "Gateway": "",
                        "DNS": "",
                    }
                ],
            },
            {
                "key": "adapters",
                "label": "Adapters",
                "success": True,
                "data": [{"Name": "Ethernet", "Status": "Disconnected"}],
            },
        ],
    )

    findings = PlaybookAnalyzer().analyze(execution)
    ids = {item.id for item in findings}
    assert "NETWORK_NO_IPV4" in ids
    assert "NETWORK_NO_GATEWAY" in ids
    assert "NETWORK_NO_DNS" in ids
    assert "NETWORK_NO_UP_ADAPTER" in ids


def test_glpi_playbook_understands_module_status_shape():
    execution = PlaybookExecution(
        playbook="glpi",
        host="PC01",
        started_at="start",
        finished_at="finish",
        steps=[
            {
                "key": "glpi",
                "label": "GLPI",
                "success": True,
                "data": {
                    "Installed": True,
                    "Running": False,
                    "Services": [],
                },
            }
        ],
    )
    findings = PlaybookAnalyzer().analyze(execution)
    assert [item.id for item in findings] == ["GLPI_STOPPED"]
