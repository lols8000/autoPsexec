from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any


class EvaluationState(str, Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    UNKNOWN = "UNKNOWN"
    NOT_APPLICABLE = "NOT_APPLICABLE"


@dataclass(frozen=True, slots=True)
class EvaluationPolicy:
    min_disk_free_percent: float = 15.0
    critical_disk_free_percent: float = 5.0
    max_uptime_days: int = 30
    defender_required: bool = True
    firewall_required: bool = True
    glpi_required: bool = True
    pending_reboot_not_allowed: bool = True
    bitlocker_required: bool = False
    tpm_required: bool = False
    secure_boot_required: bool = False

    @classmethod
    def from_mapping(cls, values: dict[str, Any] | None) -> "EvaluationPolicy":
        data = values or {}
        defaults = cls()
        return cls(
            min_disk_free_percent=float(
                data.get("min_disk_free_percent", defaults.min_disk_free_percent)
            ),
            critical_disk_free_percent=float(
                data.get(
                    "critical_disk_free_percent",
                    defaults.critical_disk_free_percent,
                )
            ),
            max_uptime_days=int(
                data.get("max_uptime_days", defaults.max_uptime_days)
            ),
            defender_required=bool(data.get("defender_required", True)),
            firewall_required=bool(data.get("firewall_required", True)),
            glpi_required=bool(data.get("glpi_required", True)),
            pending_reboot_not_allowed=bool(
                data.get("pending_reboot_not_allowed", True)
            ),
            bitlocker_required=bool(data.get("bitlocker_required", False)),
            tpm_required=bool(data.get("tpm_required", False)),
            secure_boot_required=bool(
                data.get("secure_boot_required", False)
            ),
        )


@dataclass(frozen=True, slots=True)
class CheckResult:
    key: str
    label: str
    state: EvaluationState
    severity: str
    actual: Any
    expected: Any
    message: str
    penalty: int = 0

    @property
    def compliant(self) -> bool | None:
        if self.state is EvaluationState.PASS:
            return True
        if self.state is EvaluationState.FAIL:
            return False
        return None


def number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def integer(value: Any) -> int | None:
    parsed = number(value)
    return None if parsed is None else int(parsed)


def _unknown(
    key: str,
    label: str,
    severity: str,
    expected: Any,
) -> CheckResult:
    return CheckResult(
        key,
        label,
        EvaluationState.UNKNOWN,
        severity,
        None,
        expected,
        "Não foi possível obter a métrica.",
    )


def evaluate_snapshot(
    snapshot: dict[str, Any],
    policy: EvaluationPolicy,
) -> list[CheckResult]:
    checks: list[CheckResult] = []

    disk = number(snapshot.get("DiskFreePercent"))
    if disk is None:
        checks.append(
            _unknown(
                "disk",
                "Espaço livre em C:",
                "high",
                f">= {policy.min_disk_free_percent}%",
            )
        )
    elif disk < policy.critical_disk_free_percent:
        checks.append(
            CheckResult(
                "disk",
                "Espaço livre em C:",
                EvaluationState.FAIL,
                "critical",
                disk,
                f">= {policy.min_disk_free_percent}%",
                f"Disco C: com apenas {disk:.1f}% livre.",
                25,
            )
        )
    elif disk < policy.min_disk_free_percent:
        checks.append(
            CheckResult(
                "disk",
                "Espaço livre em C:",
                EvaluationState.FAIL,
                "high",
                disk,
                f">= {policy.min_disk_free_percent}%",
                f"Disco C: abaixo do mínimo ({disk:.1f}% livre).",
                15,
            )
        )
    else:
        checks.append(
            CheckResult(
                "disk",
                "Espaço livre em C:",
                EvaluationState.PASS,
                "high",
                disk,
                f">= {policy.min_disk_free_percent}%",
                "Espaço livre dentro do baseline.",
            )
        )

    uptime = integer(snapshot.get("UptimeDays"))
    if uptime is None:
        checks.append(
            _unknown(
                "uptime",
                "Uptime",
                "low",
                f"<= {policy.max_uptime_days} dias",
            )
        )
    else:
        failed = uptime > policy.max_uptime_days
        checks.append(
            CheckResult(
                "uptime",
                "Uptime",
                EvaluationState.FAIL if failed else EvaluationState.PASS,
                "medium" if failed else "low",
                uptime,
                f"<= {policy.max_uptime_days} dias",
                (
                    f"Uptime elevado: {uptime} dias."
                    if failed
                    else "Uptime dentro do baseline."
                ),
                8 if failed else 0,
            )
        )

    if policy.pending_reboot_not_allowed:
        pending = snapshot.get("PendingReboot")
        if pending is None:
            checks.append(
                _unknown(
                    "pending_reboot",
                    "Reinicialização pendente",
                    "medium",
                    False,
                )
            )
        else:
            failed = bool(pending)
            checks.append(
                CheckResult(
                    "pending_reboot",
                    "Reinicialização pendente",
                    EvaluationState.FAIL if failed else EvaluationState.PASS,
                    "medium",
                    bool(pending),
                    False,
                    (
                        "Reinicialização pendente."
                        if failed
                        else "Sem reinicialização pendente."
                    ),
                    8 if failed else 0,
                )
            )

    stopped = integer(snapshot.get("StoppedAutoServices"))
    if stopped is None:
        checks.append(
            _unknown("auto_services", "Serviços automáticos", "high", 0)
        )
    else:
        failed = stopped > 0
        checks.append(
            CheckResult(
                "auto_services",
                "Serviços automáticos",
                EvaluationState.FAIL if failed else EvaluationState.PASS,
                "high",
                stopped,
                0,
                (
                    f"{stopped} serviço(s) automático(s) parado(s)."
                    if failed
                    else "Serviços automáticos sem desvio detectado."
                ),
                12 if failed else 0,
            )
        )

    requirements = (
        ("defender", "Microsoft Defender", "DefenderEnabled", policy.defender_required, "critical", 25),
        ("firewall", "Firewall", "FirewallEnabled", policy.firewall_required, "high", 15),
        ("glpi", "GLPI Agent", "GlpiRunning", policy.glpi_required, "medium", 8),
        ("bitlocker", "BitLocker", "BitLockerProtected", policy.bitlocker_required, "high", 12),
        ("tpm", "TPM pronto", "TPMReady", policy.tpm_required, "high", 12),
        ("secure_boot", "Secure Boot", "SecureBoot", policy.secure_boot_required, "high", 12),
    )
    for key, label, source, required, severity, penalty in requirements:
        if not required:
            continue
        actual = snapshot.get(source)
        if actual is None:
            checks.append(_unknown(key, label, severity, True))
            continue
        failed = actual is not True
        checks.append(
            CheckResult(
                key,
                label,
                EvaluationState.FAIL if failed else EvaluationState.PASS,
                severity,
                actual,
                True,
                (
                    f"{label} não atende ao baseline."
                    if failed
                    else f"{label} conforme."
                ),
                penalty if failed else 0,
            )
        )

    return checks


def score_checks(checks: list[CheckResult]) -> dict[str, Any]:
    score = max(
        0,
        100 - sum(
            check.penalty
            for check in checks
            if check.state is EvaluationState.FAIL
        ),
    )
    passed = sum(check.state is EvaluationState.PASS for check in checks)
    failed = sum(check.state is EvaluationState.FAIL for check in checks)
    unknown = sum(check.state is EvaluationState.UNKNOWN for check in checks)
    applicable = passed + failed
    compliance_score = (
        round((passed / applicable) * 100) if applicable else None
    )
    overall = (
        EvaluationState.FAIL.value
        if failed
        else EvaluationState.UNKNOWN.value
        if unknown
        else EvaluationState.PASS.value
    )
    return {
        "health_score": score,
        "compliance_score": compliance_score,
        "passed": passed,
        "failed": failed,
        "unknown": unknown,
        "applicable": applicable,
        "overall_state": overall,
    }
