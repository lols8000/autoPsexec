from modules.compliance import evaluate_compliance
from modules.health import calculate_health_score


def healthy_snapshot():
    return {
        "DiskFreePercent": 45,
        "UptimeDays": 5,
        "PendingReboot": False,
        "StoppedAutoServices": 0,
        "DefenderEnabled": True,
        "FirewallEnabled": True,
        "GlpiRunning": True,
    }


def test_health_score_healthy_machine_is_100():
    result = calculate_health_score(healthy_snapshot())
    assert result["score"] == 100
    assert result["findings"] == []


def test_health_score_penalizes_critical_disk_and_defender():
    data = healthy_snapshot()
    data.update({"DiskFreePercent": 3, "DefenderEnabled": False})
    result = calculate_health_score(data)
    assert result["score"] == 50
    assert any(item["severity"] == "critical" for item in result["findings"])


def test_health_score_never_goes_below_zero():
    data = {
        "DiskFreePercent": 1,
        "UptimeDays": 90,
        "PendingReboot": True,
        "StoppedAutoServices": 10,
        "DefenderEnabled": False,
        "FirewallEnabled": False,
        "GlpiRunning": False,
    }
    assert calculate_health_score(data)["score"] == 0


def test_compliance_all_controls_pass():
    baseline = {
        "min_disk_free_percent": 15,
        "max_uptime_days": 30,
        "defender_required": True,
        "firewall_required": True,
        "glpi_required": True,
        "pending_reboot_not_allowed": True,
    }
    report = evaluate_compliance(healthy_snapshot(), baseline)
    assert report["score"] == 100
    assert report["compliant"] == report["total"]


def test_compliance_detects_multiple_deviations():
    baseline = {
        "min_disk_free_percent": 20,
        "max_uptime_days": 15,
        "defender_required": True,
        "firewall_required": True,
        "glpi_required": True,
        "pending_reboot_not_allowed": True,
    }
    data = healthy_snapshot()
    data.update({
        "DiskFreePercent": 10,
        "UptimeDays": 45,
        "DefenderEnabled": False,
        "PendingReboot": True,
    })
    report = evaluate_compliance(data, baseline)
    failed = [item["key"] for item in report["items"] if not item["compliant"]]
    assert {"disk", "uptime", "defender", "pending_reboot"}.issubset(set(failed))
    assert report["score"] < 100
