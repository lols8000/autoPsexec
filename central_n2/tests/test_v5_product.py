import json

from core.config import ConfigLoader
from core.progress import ProgressParser
from core.updater import UpdateManager
from modules.compliance import evaluate_compliance
from core.result import CommandResult
from remediation import RemediationEngine, RemediationSpec
from reports import ReportExporter
def test_progress_parser_reads_dism_percent():
    assert ProgressParser.parse("[================  62.5% ]").percent==62.5
def test_version_comparison():
    u=UpdateManager("owner/repo","5.0.0");assert u._version_tuple("v5.1.0")>u._version_tuple("5.0.9")
def test_compliance_extended_controls():
    snap={"DiskFreePercent":50,"DefenderEnabled":True,"FirewallEnabled":True,"GlpiRunning":True,"UptimeDays":1,"PendingReboot":False,"BitLockerProtected":True,"TPMReady":True,"SecureBoot":True};base={"bitlocker_required":True,"tpm_required":True,"secure_boot_required":True};r=evaluate_compliance(snap,base)
    assert r["score"]==100 and {"bitlocker","tpm","secure_boot"}.issubset({x["key"] for x in r["items"]})


def test_remediation_engine_collects_before_and_after():
    states = iter([
        {"DiskFreePercent": 5},
        {"DiskFreePercent": 20},
    ])

    def snapshot(_host):
        return next(states)

    def action(host):
        return CommandResult(True, "cleanup", host, stdout="ok")

    result = RemediationEngine().execute(
        "PC01",
        RemediationSpec("cleanup", "Limpeza", "médio"),
        action,
        snapshotter=snapshot,
    )
    assert result.command_result.success is True
    assert result.before["DiskFreePercent"] == 5
    assert result.after["DiskFreePercent"] == 20


def test_config_loader_applies_local_override(tmp_path):
    config_dir = tmp_path / "config"
    config_dir.mkdir()
    settings = config_dir / "settings.json"
    local = config_dir / "settings.local.json"
    settings.write_text(
        json.dumps({"runtime": {"max_workers": 4}, "glpi_api": {"enabled": False}}),
        encoding="utf-8",
    )
    local.write_text(
        json.dumps({"runtime": {"max_workers": 8}, "glpi_api": {"enabled": True}}),
        encoding="utf-8",
    )

    loaded = ConfigLoader(settings).settings
    assert loaded["runtime"]["max_workers"] == 8
    assert loaded["glpi_api"]["enabled"] is True


def test_markdown_report_contains_correlation_id(tmp_path):
    report = {
        "correlation_id": "ATD123",
        "host": "PC01",
        "user": "usuario",
        "generated_at": "2026-09-03T18:00:00-04:00",
        "problem": "teste",
        "diagnosis": "ok",
        "actions": [],
        "validation": "ok",
        "result": "ok",
    }
    path = ReportExporter(tmp_path).export(report, fmt="markdown", stem="r")
    text = path.read_text(encoding="utf-8")
    assert "ATD123" in text
