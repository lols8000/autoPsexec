from core.progress import ProgressParser
from core.updater import UpdateManager
from modules.compliance import evaluate_compliance
from core.result import CommandResult
from remediation import RemediationEngine, RemediationSpec
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
