from __future__ import annotations

from typing import Any

from core.evaluation import (
    EvaluationPolicy,
    EvaluationState,
    evaluate_snapshot,
    number,
)
from .models import Finding, Severity


_SEVERITY = {
    "info": Severity.INFO,
    "low": Severity.LOW,
    "medium": Severity.MEDIUM,
    "high": Severity.HIGH,
    "critical": Severity.CRITICAL,
}


class DiagnosticEngine:
    """Converte telemetria em fatos sem inventar valores ausentes."""

    def evaluate(
        self,
        data: dict[str, Any],
        policy: dict[str, Any] | EvaluationPolicy | None = None,
    ) -> list[Finding]:
        normalized = dict(data)
        if (
            normalized.get("DiskFreePercent") is None
            and normalized.get("FreePercent") is not None
        ):
            normalized["DiskFreePercent"] = normalized["FreePercent"]

        effective = (
            policy
            if isinstance(policy, EvaluationPolicy)
            else EvaluationPolicy.from_mapping(policy)
        )

        findings: list[Finding] = []
        shared_ids = {
            "disk": (
                "DISK_LOW",
                ["Analisar perfis e temporários"],
            ),
            "uptime": (
                "UPTIME_HIGH",
                ["Avaliar reinicialização em janela autorizada"],
            ),
            "pending_reboot": (
                "PENDING_REBOOT",
                ["Validar janela e reiniciar"],
            ),
            "auto_services": (
                "AUTO_SERVICES_STOPPED",
                ["Inspecionar serviços e eventos"],
            ),
            "defender": (
                "DEFENDER_DISABLED",
                ["Validar política/GPO e estado do antivírus"],
            ),
            "firewall": (
                "FIREWALL_DISABLED",
                ["Validar política/GPO do Firewall"],
            ),
            "glpi": (
                "GLPI_STOPPED",
                ["Validar serviço e log do agente"],
            ),
            "bitlocker": (
                "BITLOCKER_NONCOMPLIANT",
                ["Validar política de criptografia"],
            ),
            "tpm": (
                "TPM_NONCOMPLIANT",
                ["Validar TPM no firmware e Windows"],
            ),
            "secure_boot": (
                "SECURE_BOOT_NONCOMPLIANT",
                ["Validar Secure Boot/UEFI"],
            ),
        }

        for check in evaluate_snapshot(normalized, effective):
            if check.state is not EvaluationState.FAIL:
                continue
            finding_id, recommendations = shared_ids[check.key]
            if (
                check.key == "disk"
                and number(check.actual) is not None
                and float(check.actual) < effective.critical_disk_free_percent
            ):
                finding_id = "DISK_CRITICAL"
                recommendations = [
                    "Executar análise de espaço",
                    "Executar limpeza segura",
                ]
            findings.append(
                Finding(
                    finding_id,
                    _SEVERITY.get(check.severity, Severity.MEDIUM),
                    check.message,
                    {
                        "key": check.key,
                        "actual": check.actual,
                        "expected": check.expected,
                    },
                    recommendations,
                )
            )

        ram = self._num(
            normalized,
            "RAMUsedPercent",
            "ram_percent",
            "MemoryPercent",
        )
        cpu = self._num(
            normalized,
            "CPUPercent",
            "cpu_percent",
            "CPUAverage",
            "CpuAverage",
        )
        disk_active = self._num(
            normalized,
            "DiskActivePercent",
            "disk_active_percent",
            "DiskAverage",
        )

        if ram is not None and ram >= 90:
            findings.append(
                Finding(
                    "MEMORY_PRESSURE",
                    Severity.HIGH,
                    f"Uso de RAM em {ram:.1f}%",
                    {"ram_percent": ram},
                    ["Analisar processos por memória"],
                )
            )

        if cpu is not None and cpu >= 90:
            findings.append(
                Finding(
                    "CPU_PRESSURE",
                    Severity.HIGH,
                    f"CPU média em {cpu:.1f}%",
                    {"cpu_percent": cpu},
                    ["Analisar processos por CPU atual"],
                )
            )

        if disk_active is not None and disk_active >= 90:
            findings.append(
                Finding(
                    "DISK_SATURATION",
                    Severity.HIGH,
                    f"Disco ativo em {disk_active:.1f}%",
                    {"disk_active_percent": disk_active},
                    ["Correlacionar com espaço livre e processos"],
                )
            )

        return findings

    @staticmethod
    def _num(data: dict[str, Any], *keys: str) -> float | None:
        for key in keys:
            value = number(data.get(key))
            if value is not None:
                return value
        return None
