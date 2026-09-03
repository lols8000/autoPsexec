from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Any


@dataclass(frozen=True)
class ComplianceItem:
    key: str
    label: str
    compliant: bool
    actual: Any
    expected: Any
    severity: str = "medium"


def evaluate_compliance(snapshot: dict, baseline: dict) -> dict:
    items: list[ComplianceItem] = []

    def add(key: str, label: str, compliant: bool, actual: Any, expected: Any, severity: str = "medium") -> None:
        items.append(ComplianceItem(key, label, compliant, actual, expected, severity))

    minimum_disk = float(baseline.get("min_disk_free_percent", 15))
    add("disk", "Espaço livre em C:", float(snapshot.get("DiskFreePercent") or 0) >= minimum_disk,
        snapshot.get("DiskFreePercent"), f">= {minimum_disk}%", "high")

    if baseline.get("defender_required", True):
        add("defender", "Microsoft Defender", snapshot.get("DefenderEnabled") is True,
            snapshot.get("DefenderEnabled"), True, "critical")

    if baseline.get("firewall_required", True):
        add("firewall", "Firewall", snapshot.get("FirewallEnabled") is True,
            snapshot.get("FirewallEnabled"), True, "high")

    if baseline.get("glpi_required", True):
        add("glpi", "GLPI Agent", snapshot.get("GlpiRunning") is True,
            snapshot.get("GlpiRunning"), True, "medium")

    max_uptime = int(baseline.get("max_uptime_days", 30))
    add("uptime", "Uptime", int(snapshot.get("UptimeDays") or 0) <= max_uptime,
        snapshot.get("UptimeDays"), f"<= {max_uptime} dias", "low")

    if baseline.get("pending_reboot_not_allowed", True):
        add("pending_reboot", "Reinicialização pendente", snapshot.get("PendingReboot") is not True,
            snapshot.get("PendingReboot"), False, "medium")

    compliant = sum(1 for item in items if item.compliant)
    score = round((compliant / len(items)) * 100) if items else 100
    return {
        "score": score,
        "compliant": compliant,
        "total": len(items),
        "items": [asdict(item) for item in items],
    }
