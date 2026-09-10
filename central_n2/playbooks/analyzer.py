from __future__ import annotations

from typing import Any

from diagnostics.engine import DiagnosticEngine
from diagnostics.models import Finding, Severity
from .base import PlaybookExecution


class PlaybookAnalyzer:
    """Interpreta evidências específicas de cada playbook de forma conservadora."""

    def __init__(self, diagnostic_engine: DiagnosticEngine | None = None) -> None:
        self.engine = diagnostic_engine or DiagnosticEngine()

    @staticmethod
    def _step_data(execution: PlaybookExecution, key: str) -> Any:
        for step in execution.steps:
            if step.get("key") == key and step.get("success"):
                return step.get("data")
        return None

    @staticmethod
    def _as_list(value: Any) -> list[Any]:
        if value is None:
            return []
        return value if isinstance(value, list) else [value]

    def analyze(
        self,
        execution: PlaybookExecution,
        *,
        policy: dict[str, Any] | None = None,
    ) -> list[Finding]:
        method = getattr(
            self,
            f"_analyze_{execution.playbook}",
            self._analyze_generic,
        )
        return method(execution, policy=policy)

    def _analyze_generic(
        self,
        execution: PlaybookExecution,
        *,
        policy: dict[str, Any] | None = None,
    ) -> list[Finding]:
        merged: dict[str, Any] = {}
        for step in execution.steps:
            data = step.get("data")
            if step.get("success") and isinstance(data, dict):
                merged.update(data)
        return self.engine.evaluate(merged, policy)

    def _analyze_slow(self, execution, *, policy=None):
        return self._analyze_generic(execution, policy=policy)

    def _analyze_disk(self, execution, *, policy=None):
        data = self._step_data(execution, "disk")
        return self.engine.evaluate(
            data if isinstance(data, dict) else {},
            policy,
        )

    def _analyze_network(self, execution, *, policy=None):
        findings: list[Finding] = []
        configs = self._as_list(self._step_data(execution, "network"))
        adapters = self._as_list(self._step_data(execution, "adapters"))

        usable_configs = [item for item in configs if isinstance(item, dict)]
        if usable_configs and not any(item.get("IPv4") for item in usable_configs):
            findings.append(
                Finding(
                    "NETWORK_NO_IPV4",
                    Severity.HIGH,
                    "Nenhuma configuração IPv4 foi encontrada.",
                    recommendations=[
                        "Validar DHCP, endereço estático e estado do adaptador"
                    ],
                )
            )
        if usable_configs and not any(item.get("Gateway") for item in usable_configs):
            findings.append(
                Finding(
                    "NETWORK_NO_GATEWAY",
                    Severity.HIGH,
                    "Nenhum gateway IPv4 foi encontrado.",
                    recommendations=["Validar DHCP/rota padrão"],
                )
            )
        if usable_configs and not any(item.get("DNS") for item in usable_configs):
            findings.append(
                Finding(
                    "NETWORK_NO_DNS",
                    Severity.MEDIUM,
                    "Nenhum servidor DNS foi encontrado.",
                    recommendations=["Validar configuração DNS do adaptador"],
                )
            )

        usable_adapters = [item for item in adapters if isinstance(item, dict)]
        if usable_adapters and not any(
            str(item.get("Status", "")).casefold() == "up"
            for item in usable_adapters
        ):
            findings.append(
                Finding(
                    "NETWORK_NO_UP_ADAPTER",
                    Severity.HIGH,
                    "Nenhum adaptador de rede está no estado Up.",
                    recommendations=["Validar cabo, Wi-Fi, driver e estado do adaptador"],
                )
            )
        return findings

    def _analyze_printer(self, execution, *, policy=None):
        findings: list[Finding] = []
        printers = self._as_list(self._step_data(execution, "printers"))
        queue = self._as_list(self._step_data(execution, "print_queue"))
        services = self._as_list(self._step_data(execution, "services"))

        if not [item for item in printers if isinstance(item, dict)]:
            findings.append(
                Finding(
                    "PRINTER_NONE",
                    Severity.MEDIUM,
                    "Nenhuma impressora foi encontrada.",
                    recommendations=["Validar instalação/mapeamento da impressora"],
                )
            )

        jobs = [item for item in queue if isinstance(item, dict)]
        if jobs:
            findings.append(
                Finding(
                    "PRINT_QUEUE_PENDING",
                    Severity.MEDIUM,
                    f"{len(jobs)} trabalho(s) encontrado(s) na fila de impressão.",
                    {"count": len(jobs)},
                    ["Validar status dos jobs e Spooler antes de limpar a fila"],
                )
            )

        spooler = next(
            (
                item
                for item in services
                if isinstance(item, dict)
                and str(item.get("Name", "")).casefold() == "spooler"
            ),
            None,
        )
        if spooler and str(spooler.get("Status", "")).casefold() != "running":
            findings.append(
                Finding(
                    "SPOOLER_STOPPED",
                    Severity.HIGH,
                    "Serviço Spooler não está em execução.",
                    spooler,
                    ["Reiniciar o Spooler e validar novamente"],
                )
            )
        return findings

    def _analyze_domain(self, execution, *, policy=None):
        data = self._step_data(execution, "domain")
        if not isinstance(data, dict):
            return []

        findings: list[Finding] = []
        if data.get("PartOfDomain") is False:
            findings.append(
                Finding(
                    "DOMAIN_NOT_JOINED",
                    Severity.MEDIUM,
                    "A estação não está ingressada em domínio.",
                )
            )
            return findings

        if data.get("SecureChannel") is False:
            findings.append(
                Finding(
                    "DOMAIN_SECURE_CHANNEL_BROKEN",
                    Severity.HIGH,
                    "O secure channel do domínio falhou.",
                    recommendations=[
                        "Validar hora/DC antes de reparar o secure channel"
                    ],
                )
            )
        if data.get("PartOfDomain") is True and not data.get("DomainController"):
            findings.append(
                Finding(
                    "DOMAIN_CONTROLLER_UNRESOLVED",
                    Severity.MEDIUM,
                    "Não foi possível identificar um controlador de domínio.",
                    recommendations=["Validar DNS, conectividade com DC e horário"],
                )
            )
        return findings

    def _analyze_update(self, execution, *, policy=None):
        data = self._step_data(execution, "updates")
        if not isinstance(data, dict):
            return []

        wu = data.get("WindowsUpdate")
        if not isinstance(wu, dict):
            return []

        findings: list[Finding] = []
        if wu.get("Error"):
            findings.append(
                Finding(
                    "UPDATE_QUERY_FAILED",
                    Severity.MEDIUM,
                    "A consulta do Windows Update falhou.",
                    {"error": wu.get("Error")},
                    ["Validar serviços e componentes do Windows Update"],
                )
            )
            return findings

        count = wu.get("PendingCount")
        if isinstance(count, int) and count > 0:
            findings.append(
                Finding(
                    "UPDATE_PENDING",
                    Severity.INFO,
                    f"{count} atualização(ões) pendente(s).",
                    {"count": count},
                )
            )
        pending = self._as_list(wu.get("Pending"))
        if any(
            isinstance(item, dict) and item.get("RebootRequired")
            for item in pending
        ):
            findings.append(
                Finding(
                    "UPDATE_REBOOT_REQUIRED",
                    Severity.MEDIUM,
                    "Há atualização pendente que requer reinicialização.",
                    recommendations=["Validar janela de reinicialização"],
                )
            )
        return findings

    def _analyze_crash(self, execution, *, policy=None):
        crashes = [
            item
            for item in self._as_list(
                self._step_data(execution, "app_crashes")
            )
            if isinstance(item, dict)
        ]
        if not crashes:
            return []
        return [
            Finding(
                "APP_CRASHES_RECENT",
                Severity.MEDIUM,
                f"{len(crashes)} evento(s) recente(s) de crash foram encontrados.",
                {"count": len(crashes)},
                ["Correlacionar aplicativo, módulo com falha e horário"],
            )
        ]

    def _analyze_bsod(self, execution, *, policy=None):
        data = self._step_data(execution, "bsod")
        if not isinstance(data, dict):
            return []
        events = self._as_list(data.get("BugcheckEvents"))
        dumps = self._as_list(data.get("Minidumps"))
        memory_dump = data.get("MemoryDump")
        if not events and not dumps and not memory_dump:
            return []
        return [
            Finding(
                "BSOD_EVIDENCE",
                Severity.HIGH,
                "Foram encontradas evidências de tela azul/bugcheck.",
                {
                    "events": len(events),
                    "minidumps": len(dumps),
                    "memory_dump": bool(memory_dump),
                },
                ["Preservar dumps e correlacionar código de bugcheck/driver"],
            )
        ]

    def _analyze_glpi(self, execution, *, policy=None):
        data = self._step_data(execution, "glpi")
        if not isinstance(data, dict):
            return []
        if data.get("Installed") is False:
            return [
                Finding(
                    "GLPI_NOT_INSTALLED",
                    Severity.MEDIUM,
                    "GLPI Agent não foi encontrado.",
                    recommendations=["Validar política de instalação do agente"],
                )
            ]
        if data.get("Running") is False:
            return [
                Finding(
                    "GLPI_STOPPED",
                    Severity.MEDIUM,
                    "GLPI Agent está instalado, mas não está em execução.",
                    recommendations=["Consultar log e validar serviço"],
                )
            ]
        return []
