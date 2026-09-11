from __future__ import annotations

import getpass
import json
import re
import time
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from core.actions import ActionSpec
from core.baselines import BaselineRepository
from core.config import ConfigLoader
from core.context import AttendanceContext
from core.jobs import JobManager, OperationClass
from core.result import CommandResult
from core.session import SessionManager
from core.updater import UpdateManager
from core.validation import validate_host
from core.version import __version__
from diagnostics.correlation import CorrelationEngine
from diagnostics.engine import DiagnosticEngine
from execution import (
    ExecutionBlockedError,
    ExecutionDependencies,
    ExecutionEngine,
    ExecutionPolicyContext,
    ParameterKind,
    PolicyState,
    RecoveryResult,
    RiskLevel,
    SelectorKind,
    build_execution_registry,
)
from integrations.glpi.client import GLPIClient, GLPIError
from modules.compliance import evaluate_compliance
from modules.health import calculate_health_score
from playbooks import PlaybookAnalyzer, PlaybookRunner, builtin_playbooks
from remediation import (
    RemediationEngine,
    RemediationSpec,
    ValidationStatus,
    validate_cleanup,
    validate_gpupdate,
    validate_spooler,
    validate_windows_update_reset,
)
from reports import ReportExporter, SupportReportBuilder
from storage import CentralDatabase, diff_values
from ui.console_base import ConsoleBase


class ConsoleUIV5(ConsoleBase):
    """Interface principal da Central N2 v5."""

    MENU_LINES = (
        "[1] Selecionar estação",
        "[2] Saúde / Compliance",
        "[3] Performance",
        "[4] Reparo do Windows",
        "[5] Hardware / Drivers / Dispositivos",
        "[6] Inicialização / Tarefas",
        "[7] Crashes / BSOD",
        "[8] Segurança",
        "[9] Rede",
        "[10] Usuários / Perfis",
        "[11] Software / GLPI Agent",
        "[12] Impressoras",
        "[13] Domínio / GPO",
        "[14] Disco / Armazenamento / Bateria",
        "[15] Ferramentas avançadas",
        "[16] Sysinternals",
        "[17] Pacote de diagnóstico",
        "[18] Energia / Processos / Serviços",
        "[19] Conectividade / Capabilities",
        "[20] Assistente N2 / Playbooks",
        "[21] Histórico / Diff",
        "[22] Gerar relatório",
        "[23] Jobs",
        "[24] Atualização da Central",
        "[25] GLPI API",
        "[26] Remediações guiadas",
        "[27] Perfil / Baseline",
        "[28] Central de Execuções",
        "[0] Sair",
    )

    def __init__(
        self,
        executor,
        settings_path: Path,
        *,
        settings: dict[str, Any] | None = None,
    ) -> None:
        settings = settings or ConfigLoader(settings_path).settings
        runtime = settings.get("runtime", {})
        ui = settings.get("ui", {})

        self.job_manager = JobManager(
            max_workers=int(runtime.get("max_workers", 6)),
            heartbeat_seconds=float(ui.get("heartbeat_seconds", 0.2)),
        )
        super().__init__(
            executor,
            settings_path,
            settings=settings,
            job_manager=self.job_manager,
        )

        root = settings_path.parent.parent
        self.sessions = SessionManager(executor)
        self.baselines = BaselineRepository(root / "baselines")
        self.active_baseline_profile = str(
            settings.get("compliance", {}).get("profile", "DEFAULT")
        ).upper()

        persistence = settings.get("persistence", {})
        db_path = Path(persistence.get("database", "data/central_n2.db"))
        if not db_path.is_absolute():
            db_path = root / db_path

        self.db: CentralDatabase | None = None
        if persistence.get("enabled", True):
            self.db = CentralDatabase(db_path)
            if not self.db.integrity_check():
                raise RuntimeError(
                    "O banco SQLite da Central N2 falhou no integrity check."
                )
            self.db.prune(
                int(persistence.get("snapshot_retention_days", 180))
            )
            self.job_manager.set_observer(self.db.save_job)

        self.engine = DiagnosticEngine()
        self.correlator = CorrelationEngine()
        self.playbook_runner = PlaybookRunner()
        self.playbook_analyzer = PlaybookAnalyzer(self.engine)
        self.playbooks = builtin_playbooks()
        self.remediation_engine = RemediationEngine()
        self.execution_registry = build_execution_registry(
            ExecutionDependencies(
                system=self.system,
                network=self.network,
                software=self.software,
                printers=self.printers,
                devices=self.devices,
                domain=self.domain,
                users=self.users,
                disk=self.disk,
                glpi=self.glpi,
                health=self.health,
                security=self.security,
                updates=self.updates,
                repair=self.repair,
                packages=self.packages,
                certificates=self.certificates,
                registry_actions=self.registry_actions,
                files=self.file_ops,
            )
        )
        self.execution_engine = ExecutionEngine(
            self.execution_registry,
            remediation_engine=self.remediation_engine,
        )
        self.report_builder = SupportReportBuilder()
        self.report_exporter = ReportExporter(root / "reports" / "support")

        update_config = settings.get("updates", {})
        self.updates_enabled = bool(
            update_config.get("enabled", True)
        )
        self.updater = UpdateManager(
            update_config.get("repository", "lols8000/autoPsexec"),
            __version__,
        )
        self.update_dir = root / "updates"

        self.context = AttendanceContext.start()

    def _handlers(self) -> dict[str, Callable[[], None]]:
        return {
            "1": self.select_host,
            "2": self.menu_health,
            "3": self.menu_performance,
            "4": self.menu_repair,
            "5": self.menu_devices,
            "6": self.menu_startup_tasks,
            "7": self.menu_crashes,
            "8": self.menu_security,
            "9": self.menu_network,
            "10": self.menu_users,
            "11": self.menu_software_glpi,
            "12": self.menu_printers,
            "13": self.menu_domain,
            "14": self.menu_storage,
            "15": self.menu_tools,
            "16": self.menu_sysinternals,
            "17": self.collect_diagnostic,
            "18": self.menu_system,
            "19": self.menu_connectivity,
            "20": self.menu_playbooks,
            "21": self.menu_history,
            "22": self.menu_report,
            "23": self.menu_jobs,
            "24": self.menu_update,
            "25": self.menu_glpi_api,
            "26": self.menu_remediations,
            "27": self.menu_baseline,
            "28": self.menu_execution,
        }

    def run(self) -> None:
        try:
            while True:
                self.clear()
                print("╔════════════════════════════════════════════════════╗")
                print(
                    f"║ CENTRAL N2 WORKSTATION — V{__version__:<24}║"
                )
                print("╚════════════════════════════════════════════════════╝")

                transport = (
                    self.context.session.transport
                    if self.context.session
                    else "-"
                )
                state = (
                    self.context.session.connectivity.get("state", "-")
                    if self.context.session
                    else "-"
                )
                print(
                    f"\nAlvo: {self.host or 'nenhum'} | "
                    f"Transporte: {transport} | Estado: {state}"
                )

                for line in self.MENU_LINES:
                    print(line)

                option = input("\nOpção: ").strip()
                if option == "0":
                    return

                handler = self._handlers().get(option)
                if not handler:
                    continue

                try:
                    handler()
                except KeyboardInterrupt:
                    print("\nOperação cancelada pelo operador.")
                    self.pause()
                except Exception as exc:
                    self._handle_ui_error(option, exc)
        finally:
            self.jobs.shutdown()
            self.job_manager.shutdown()

    def _handle_ui_error(self, option: str, exc: Exception) -> None:
        logger = getattr(self.executor, "logger", None)
        if logger:
            with logger.bind(correlation_id=self.context.correlation_id):
                logger.log_event(
                    "ui_error",
                    self.host or "local-ui",
                    "failure",
                    menu_option=option,
                    error=f"{type(exc).__name__}: {exc}",
                )

        print(
            f"\n✗ ERRO INTERNO: {type(exc).__name__}: {exc}"
        )
        print(f"Correlation ID: {self.context.correlation_id}")
        print(
            "A operação foi encerrada, mas a Central continua disponível."
        )
        self.pause()

    def _reset_attendance(self, host: str) -> None:
        self.context = AttendanceContext.start(host)

    def select_host(self) -> None:
        self.clear()
        try:
            host = validate_host(
                input("Hostname ou IP da estação: ").strip()
            )
        except ValueError as exc:
            print(f"Entrada inválida: {exc}")
            self.pause()
            return

        self._reset_attendance(host)
        self.host = host

        try:
            session = self.jobs.run(
                "Abrindo sessão lógica",
                self._trace(
                    lambda: self.sessions.open(host, refresh=True),
                    action="session.preflight",
                ),
                timeout=180,
                operation_class=OperationClass.READ_ONLY,
                host=host,
                correlation_id=self.context.correlation_id,
            )
        except Exception as exc:
            print(f"\n✗ Falha no preflight: {exc}")
            self.pause()
            return

        self.context.session = session

        connectivity = session.connectivity
        print(f"\n✓ Transporte selecionado: {session.transport}")
        print(
            "Local: "
            f"{'SIM' if connectivity.get('is_local') else 'NÃO'} | "
            f"DNS: {connectivity.get('dns')} | "
            f"445: {connectivity.get('tcp_445')} | "
            f"5985: {connectivity.get('tcp_5985')} | "
            f"WinRM: {connectivity.get('winrm')} | "
            f"ADMIN$: {connectivity.get('admin_share')} | "
            f"PsExec: {connectivity.get('psexec')}"
        )
        print(
            f"Estado: {connectivity.get('state')} | "
            f"Diagnóstico: {connectivity.get('diagnosis')}"
        )

        if not session.ready:
            print(
                "\n⚠ Nenhum transporte administrativo foi validado. "
                "Use o menu 19 para revisar conectividade."
            )
            self.pause()
            return

        result = self.execute(
            ActionSpec.read(
                "health.initial_snapshot",
                "Snapshot inicial de saúde",
                timeout_seconds=120,
            ),
            lambda: self.health.snapshot(host),
        )
        if result and result.success and isinstance(result.data, dict):
            self._set_health_snapshot(result.data)
            self.show_health(result.data)
            self._persist_health(result.data)

        self.pause()

    def _baseline(self) -> dict[str, Any]:
        return self.baselines.resolve(
            self.active_baseline_profile,
            self.settings.get("compliance", {}),
        )

    @staticmethod
    def _metric(value: Any, suffix: str = "") -> str:
        return "-" if value is None else f"{value}{suffix}"

    def show_health(self, snapshot: dict[str, Any]) -> None:
        health = calculate_health_score(
            snapshot,
            baseline=self._baseline(),
        )

        print("\n=== SAÚDE DA ESTAÇÃO ===")
        print(
            f"Score: {health['score']}/100 | "
            f"Estado: {health['overall_state']} | "
            f"CPU: {self._metric(snapshot.get('CPUPercent'), '%')} | "
            f"RAM: {self._metric(snapshot.get('RAMUsedPercent'), '%')} | "
            f"Disco livre: "
            f"{self._metric(snapshot.get('DiskFreePercent'), '%')} | "
            f"Uptime: {self._metric(snapshot.get('UptimeDays'), ' dias')}"
        )

        if health["unknown"]:
            print(f"⚠ Métricas indisponíveis: {health['unknown']}")

        for item in health["findings"]:
            symbol = "?" if item.get("state") == "UNKNOWN" else "!"
            print(
                f" - [{symbol} {item['severity'].upper()}] "
                f"{item['message']}"
            )

        if not health["findings"]:
            print("Nenhum desvio relevante encontrado.")

    def _trace(
        self,
        func: Callable,
        **context: Any,
    ) -> Callable:
        def wrapped():
            logger = getattr(self.executor, "logger", None)
            if not logger:
                return func()
            with logger.bind(
                correlation_id=self.context.correlation_id,
                **context,
            ):
                return func()

        return wrapped

    def _set_health_snapshot(self, data: dict[str, Any]) -> None:
        self.context.health_snapshot = data

    def _persist_health(self, data: dict[str, Any]) -> None:
        findings = self.engine.evaluate(data, self._baseline())
        diagnoses = self.correlator.correlate(findings)
        self.context.diagnoses = list(diagnoses)

        if not self.db:
            return

        self.db.save_snapshot(
            self.host,
            data,
            kind="health",
            correlation_id=self.context.correlation_id,
        )
        for finding in findings:
            self.db.save_finding(
                self.host,
                finding.id,
                finding.severity.value,
                asdict(finding),
                correlation_id=self.context.correlation_id,
            )

    def menu_health(self) -> None:
        if not self.require_host():
            return

        self.clear()
        result = self.execute(
            ActionSpec.read(
                "health.collect",
                "Coletando saúde e compliance",
                timeout_seconds=120,
            ),
            lambda: self.health.snapshot(self.host),
        )

        if (
            result
            and result.success
            and isinstance(result.data, dict)
        ):
            self._set_health_snapshot(result.data)
            self.show_health(result.data)
            self._persist_health(result.data)

            compliance = evaluate_compliance(
                result.data,
                self._baseline(),
            )
            score = (
                "-"
                if compliance["score"] is None
                else compliance["score"]
            )
            print(
                f"\nCompliance: {score}/100 | "
                f"estado {compliance['overall_state']} | "
                f"PASS {compliance['compliant']} | "
                f"FAIL {compliance['failed']} | "
                f"UNKNOWN {compliance['unknown']} | "
                f"N/A {compliance['not_applicable']}"
            )

            symbols = {
                "PASS": "✓",
                "FAIL": "✗",
                "UNKNOWN": "?",
                "NOT_APPLICABLE": "-",
            }
            for item in compliance["items"]:
                print(
                    f" {symbols.get(item['state'], '?')} "
                    f"{item['label']}: {item['actual']} | "
                    f"esperado {item['expected']} | {item['state']}"
                )

            for diagnosis in self.context.diagnoses:
                print(
                    f" - {diagnosis.title} | "
                    f"confiança {diagnosis.confidence}: "
                    f"{diagnosis.rationale}"
                )

        self.pause()

    def menu_repair(self) -> None:
        if not self.require_host():
            return

        self.clear()
        print(
            "REPARO DO WINDOWS\n"
            "1 - SFC\n"
            "2 - DISM CheckHealth\n"
            "3 - DISM ScanHealth\n"
            "4 - DISM RestoreHealth (progresso real)\n"
            "5 - Analisar Component Store\n"
            "6 - Limpar Component Store\n"
            "7 - CHKDSK online\n"
            "8 - Verificar WMI\n"
            "0 - Voltar"
        )
        option = input("Opção: ").strip()

        if option == "4":
            if not self.confirm(
                f"Executar DISM RestoreHealth em {self.host}?"
            ):
                return

            def progress(event) -> None:
                if event.percent is not None:
                    print(
                        f"\rDISM RestoreHealth "
                        f"[{event.percent:6.2f}%] "
                        f"{event.message[:60]:60}",
                        end="",
                        flush=True,
                    )

            try:
                result = self.jobs.run(
                    "DISM RestoreHealth",
                    self._trace(
                        lambda: self.repair.dism_restorehealth_progress(
                            self.host,
                            progress,
                        ),
                        action="repair.dism.restore",
                    ),
                    operation_class=OperationClass.HEAVY_WRITE,
                    timeout=3600,
                    on_tick=lambda _: None,
                    host=self.host,
                    correlation_id=self.context.correlation_id,
                )
                print()
                self.show_result(result)
            except Exception as exc:
                print(f"\n✗ {type(exc).__name__}: {exc}")

            self.pause()
            return

        actions: dict[str, tuple[ActionSpec, Callable]] = {
            "1": (
                ActionSpec(
                    "repair.sfc",
                    "SFC /scannow",
                    OperationClass.HEAVY_WRITE,
                    1800,
                    True,
                ),
                self.repair.sfc_scan,
            ),
            "2": (
                ActionSpec.read(
                    "repair.dism.check",
                    "DISM CheckHealth",
                    timeout_seconds=600,
                ),
                self.repair.dism_checkhealth,
            ),
            "3": (
                ActionSpec.read(
                    "repair.dism.scan",
                    "DISM ScanHealth",
                    timeout_seconds=1800,
                ),
                self.repair.dism_scanhealth,
            ),
            "5": (
                ActionSpec.read(
                    "repair.component_store.analyze",
                    "Analisar Component Store",
                    timeout_seconds=900,
                ),
                self.repair.component_store,
            ),
            "6": (
                ActionSpec(
                    "repair.component_store.cleanup",
                    "Limpar Component Store",
                    OperationClass.HEAVY_WRITE,
                    1800,
                    True,
                ),
                self.repair.component_cleanup,
            ),
            "7": (
                ActionSpec.read(
                    "repair.chkdsk.scan",
                    "CHKDSK /scan",
                    timeout_seconds=1800,
                ),
                self.repair.chkdsk_scan,
            ),
            "8": (
                ActionSpec.read(
                    "repair.wmi.verify",
                    "Verificar WMI",
                    timeout_seconds=300,
                ),
                self.repair.repository_consistency,
            ),
        }

        item = actions.get(option)
        if not item:
            return

        spec, function = item
        if spec.requires_confirmation and not self.confirm(
            f"Executar {spec.title} em {self.host}?"
        ):
            return

        self.execute(spec, lambda: function(self.host))
        self.pause()

    def _capability(self, key: str, default=None):
        if not self.context.session:
            return default
        return self.context.session.capabilities.get(key, default)

    def menu_storage(self) -> None:
        if not self.require_host():
            return

        battery = self._capability("Battery")
        battery_label = (
            "Bateria"
            if battery is not False
            else "Bateria [N/A nesta estação]"
        )

        self.clear()
        print("DISCO / ARMAZENAMENTO / BATERIA")
        print("1 - Espaço e volumes")
        print("2 - Discos físicos / saúde")
        print(f"3 - {battery_label}")
        print("4 - Perfis por tamanho")
        print("5 - Estimar limpeza")
        print("0 - Voltar")

        option = input("Opção: ").strip()
        if option == "3" and battery is False:
            print(
                "Esta estação não expõe bateria ao Windows; "
                "a coleta foi evitada."
            )
            self.pause()
            return

        actions = {
            "1": (
                ActionSpec.read(
                    "storage.volumes",
                    "Analisando volumes",
                    timeout_seconds=600,
                ),
                self.disk.space,
            ),
            "2": (
                ActionSpec.read(
                    "storage.physical_disks",
                    "Analisando discos físicos",
                    timeout_seconds=600,
                ),
                self.storage.physical_disks,
            ),
            "3": (
                ActionSpec.read(
                    "storage.battery",
                    "Analisando bateria",
                    timeout_seconds=600,
                ),
                self.storage.battery,
            ),
            "4": (
                ActionSpec(
                    "storage.profile_sizes",
                    "Medindo perfis",
                    OperationClass.HEAVY_READ,
                    600,
                ),
                self.disk.profile_sizes,
            ),
            "5": (
                ActionSpec(
                    "storage.cleanup_estimate",
                    "Estimando limpeza",
                    OperationClass.HEAVY_READ,
                    600,
                ),
                self.disk.cleanup_estimate,
            ),
        }

        item = actions.get(option)
        if not item:
            return

        spec, function = item
        self.execute(spec, lambda: function(self.host))
        self.pause()

    def menu_software_glpi(self) -> None:
        if not self.require_host():
            return

        winget = self._capability("Winget")
        winget_label = (
            "Winget" if winget is not False else "Winget [N/A]"
        )

        self.clear()
        print("SOFTWARE / GLPI")
        print("1 - Software instalado")
        print(f"2 - {winget_label}")
        print("3 - GLPI status")
        print("4 - Forçar inventário GLPI")
        print("0 - Voltar")

        option = input("Opção: ").strip()
        if option == "2" and winget is False:
            print("Winget não foi detectado nesta estação.")
            self.pause()
            return

        if option == "1":
            spec = ActionSpec.read(
                "software.inventory",
                "Inventariando software",
                timeout_seconds=600,
            )
            function = self.software.list_installed
        elif option == "2":
            spec = ActionSpec.read(
                "software.winget",
                "Verificando Winget",
                timeout_seconds=120,
            )
            function = self.software.winget_available
        elif option == "3":
            spec = ActionSpec.read(
                "glpi.status",
                "Verificando GLPI",
                timeout_seconds=180,
            )
            function = self.glpi.status
        elif option == "4":
            spec = ActionSpec(
                "glpi.force_inventory",
                "Forçando inventário GLPI",
                OperationClass.LIGHT_WRITE,
                600,
            )
            function = self.glpi.force_inventory
        else:
            return

        self.execute(spec, lambda: function(self.host))
        self.pause()

    def menu_connectivity(self) -> None:
        if not self.require_host():
            return

        self.clear()
        try:
            session = self.jobs.run(
                "Atualizando conectividade/capabilities",
                self._trace(
                    lambda: self.sessions.open(
                        self.host,
                        refresh=True,
                    ),
                    action="connectivity.refresh",
                ),
                timeout=180,
                operation_class=OperationClass.READ_ONLY,
                host=self.host,
                correlation_id=self.context.correlation_id,
            )
            self.context.session = session
            self.context.session = session
            print(
                json.dumps(
                    {
                        "transport": session.transport,
                        "connectivity": session.connectivity,
                        "capabilities": session.capabilities,
                        "capability_error": session.capability_error,
                    },
                    indent=2,
                    ensure_ascii=False,
                    default=str,
                )
            )
        except Exception as exc:
            print(f"✗ {type(exc).__name__}: {exc}")

        self.pause()

    def _collectors(self) -> dict[str, Callable]:
        return {
            "health": self.health.snapshot,
            "performance": lambda host: self.performance.snapshot(
                host,
                8,
                1,
            ),
            "disk": self.disk.space,
            "processes": self.system.list_processes,
            "startup": self.startup.overview,
            "network": self.network.ip_configuration,
            "adapters": self.network.adapters,
            "connections": self.network.connections,
            "proxy": self.tools.proxy,
            "printers": self.printers.list_printers,
            "print_queue": self.printers.queue,
            "services": self.system.list_services,
            "domain": self.domain.status,
            "gpresult": self.domain.gpresult,
            "updates": self.updates.status,
            "app_crashes": self.crashes.app_crashes,
            "bsod": self.crashes.bsod_history,
            "devices": self.devices.problem_devices,
            "profiles": self.disk.profile_sizes,
            "cleanup_estimate": self.disk.cleanup_estimate,
            "glpi": self.glpi.status,
            "glpi_log": self.glpi.recent_log,
        }

    def menu_playbooks(self) -> None:
        if not self.require_host():
            return

        self.clear()
        keys = list(self.playbooks)
        print("ASSISTENTE N2 / PLAYBOOKS")
        for index, key in enumerate(keys, 1):
            print(f"{index} - {self.playbooks[key].title}")

        option = input("Opção (0 volta): ").strip()
        if (
            option == "0"
            or not option.isdigit()
            or not 1 <= int(option) <= len(keys)
        ):
            return

        spec = self.playbooks[keys[int(option) - 1]]
        operation_class = (
            OperationClass.HEAVY_READ
            if spec.key in {"slow", "disk"}
            else OperationClass.READ_ONLY
        )

        try:
            execution = self.jobs.run(
                f"Playbook {spec.title}",
                self._trace(
                    lambda: self.playbook_runner.run(
                        spec,
                        self.host,
                        self._collectors(),
                        on_step=lambda index, total, step: print(
                            f"[{index}/{total}] {step.label}..."
                        ),
                    ),
                    action=f"playbook.{spec.key}",
                ),
                operation_class=operation_class,
                timeout=self.long_timeout,
                host=self.host,
                correlation_id=self.context.correlation_id,
            )
        except Exception as exc:
            print(f"✗ {type(exc).__name__}: {exc}")
            self.pause()
            return

        self.context.playbook = execution

        for step in execution.steps:
            print(
                f"{'✓' if step.get('success') else '✗'} "
                f"{step['label']} "
                f"[{step.get('transport', '-')}] "
                f"{step.get('error') or ''}"
            )

        findings = self.playbook_analyzer.analyze(
            execution,
            policy=self._baseline(),
        )
        self.context.diagnoses = list(
            self.correlator.correlate(findings)
        )

        if not findings:
            print(
                "- Análise: nenhuma causa objetiva foi confirmada; "
                "resultado inconclusivo."
            )

        if self.db:
            self.db.save_snapshot(
                self.host,
                asdict(execution),
                kind=f"playbook:{spec.key}",
                correlation_id=self.context.correlation_id,
            )
            for finding in findings:
                self.db.save_finding(
                    self.host,
                    finding.id,
                    finding.severity.value,
                    asdict(finding),
                    correlation_id=self.context.correlation_id,
                )

        for diagnosis in self.context.diagnoses:
            print(
                f"- {diagnosis.title} ({diagnosis.confidence}): "
                f"{diagnosis.rationale}"
            )

        self.pause()

    def menu_history(self) -> None:
        if not self.require_host():
            return

        self.clear()
        if not self.db:
            print("Persistência desabilitada.")
            self.pause()
            return

        for record in self.db.recent_snapshots(
            self.host,
            limit=10,
        ):
            print(
                f"#{record['id']} {record['created_at']} "
                f"{record['kind']} "
                f"[{record.get('correlation_id') or '-'}]"
            )

        print("\nEXECUÇÕES RECENTES:")
        executions = self.db.recent_executions(
            self.host,
            limit=10,
        )
        if not executions:
            print("Nenhuma execução registrada.")
        for item in executions:
            rollback_marker = (
                f" rollback-of=#{item.get('rollback_of')}"
                if item.get("is_rollback")
                else ""
            )
            duration = item.get("duration_ms")
            duration_text = (
                f"{int(duration) / 1000:.1f}s"
                if duration is not None
                else "-"
            )
            print(
                f"#{item['id']} {item['created_at']} | "
                f"{item['validation_state']:<7} | "
                f"{item.get('risk') or '-':<8} | "
                f"{item.get('transport') or '-':<7} | "
                f"{duration_text:<7} | "
                f"{item['action']} | "
                f"{item.get('operator') or '-'} | "
                f"{item.get('correlation_id') or '-'}"
                f"{rollback_marker}"
            )

        print("\nDIFF DOS DOIS ÚLTIMOS HEALTH:")
        changes = self.db.diff_latest(self.host, kind="health")
        if not changes:
            print("Sem dois snapshots comparáveis.")
        for change in changes[:100]:
            print(
                f"- {change['path']}: {change['before']} -> "
                f"{change['after']} ({change['type']})"
            )

        self.pause()

    def menu_report(self) -> None:
        if not self.require_host():
            return

        self.clear()
        problem = input("Problema/resumo do chamado: ").strip()
        diagnosis = (
            "; ".join(
                f"{item.title} ({item.confidence})"
                for item in self.context.diagnoses
            )
            or "Sem diagnóstico correlacionado registrado."
        )

        actions = (
            [
                step["label"]
                for step in self.context.playbook.steps
                if step.get("success")
            ]
            if self.context.playbook
            else []
        )

        validation: Any = self.context.health_snapshot
        if self.context.execution:
            execution = self.context.execution
            actions.append(
                f"Execução: {execution.action.title} — "
                f"{execution.remediation.validation.status.value}"
            )
            validation = {
                "type": "execution",
                "action": execution.action.key,
                "action_version": execution.action.action_version,
                "risk": execution.action.risk.value,
                "transport": execution.remediation.command_result.transport,
                "operator": execution.operator,
                "duration_ms": execution.duration_ms,
                "parameters": execution.public_parameters,
                "status": execution.remediation.validation.status.value,
                "message": execution.remediation.validation.message,
                "evidence": execution.remediation.validation.evidence,
                "recovery": execution.recovery,
                "rollback_available": execution.rollback_available,
                "rollback_performed": (
                    execution.rollback_result is not None
                ),
            }
        elif self.context.remediation:
            remediation = self.context.remediation
            actions.append(
                f"Remediação: {remediation.spec.title} — "
                f"{remediation.validation.status.value}"
            )
            validation = {
                "status": remediation.validation.status.value,
                "message": remediation.validation.message,
                "evidence": remediation.validation.evidence,
                "after": remediation.after,
            }

        report = self.report_builder.build(
            host=self.host,
            user=(self.context.health_snapshot or {}).get("User"),
            problem=problem,
            diagnosis=diagnosis,
            actions=actions,
            validation=validation,
            result="Diagnóstico/atendimento registrado",
        )
        report["correlation_id"] = self.context.correlation_id

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        stem = f"{self.host}_{stamp}"
        path = self.report_exporter.export(
            report,
            fmt="markdown",
            stem=stem,
        )
        self.report_exporter.export(
            report,
            fmt="json",
            stem=stem,
        )

        self.context.report_path = path

        if self.db:
            self.db.save_report(
                self.host,
                "markdown",
                path.read_text(encoding="utf-8"),
                path=str(path),
                correlation_id=self.context.correlation_id,
            )

        print(f"✓ Relatório: {path}")
        self.pause()

    def menu_jobs(self) -> None:
        self.clear()
        print("JOBS DA SESSÃO")

        records = self.job_manager.list_records()[-30:]
        if not records:
            print("Nenhum job nesta sessão.")

        for record in records:
            print(
                f"{record.job_id} | {record.host} | "
                f"{record.state.value} | "
                f"{record.operation_class.value} | "
                f"{record.elapsed_seconds:.1f}s | "
                f"{record.label} | "
                f"{record.correlation_id or '-'}"
            )

        if self.db:
            print("\nÚLTIMOS JOBS PERSISTIDOS")
            for item in self.db.recent_jobs(
                self.host,
                limit=15,
            ):
                print(
                    f"{item['job_id']} | {item['host']} | "
                    f"{item['state']} | {item['label']} | "
                    f"{item.get('correlation_id') or '-'}"
                )

        self.pause()

    def menu_update(self) -> None:
        self.clear()
        print(f"Versão atual: {__version__}")

        if not self.updates_enabled:
            print("Atualizações estão desabilitadas pela configuração.")
            self.pause()
            return

        try:
            info = self.jobs.run(
                "Consultando releases",
                self.updater.check_latest,
                timeout=30,
                operation_class=OperationClass.READ_ONLY,
                host="github-release",
                correlation_id=self.context.correlation_id,
            )
        except Exception as exc:
            print(f"Não foi possível consultar atualização: {exc}")
            self.pause()
            return

        print(
            f"Última versão: {info.latest} | "
            f"Atualização disponível: "
            f"{'SIM' if info.update_available else 'NÃO'}"
        )
        if not info.update_available or not info.assets:
            self.pause()
            return

        print("\nArtefatos disponíveis:")
        for index, asset in enumerate(info.assets, 1):
            digest = asset.get("digest") or "sem digest publicado"
            print(
                f"{index} - {asset.get('name', 'sem nome')} | "
                f"{asset.get('size', '?')} bytes | {digest}"
            )

        choice = input(
            "Número para baixar (0 cancela): "
        ).strip()
        if (
            choice == "0"
            or not choice.isdigit()
            or not 1 <= int(choice) <= len(info.assets)
        ):
            return

        asset = info.assets[int(choice) - 1]
        if not self.confirm(
            f"Baixar {asset.get('name', 'artefato')} "
            "para a pasta updates?"
        ):
            return

        try:
            path = self.jobs.run(
                "Baixando atualização",
                lambda: self.updater.download_asset(
                    asset,
                    self.update_dir,
                ),
                timeout=900,
                operation_class=OperationClass.HEAVY_READ,
                host="github-release",
                correlation_id=self.context.correlation_id,
            )
            print(f"✓ Download verificado: {path}")
            print(
                "A instalação permanece manual/controlada; "
                "a Central não se substitui silenciosamente."
            )
        except Exception as exc:
            print(f"✗ Falha no download: {exc}")

        self.pause()

    def menu_glpi_api(self) -> None:
        self.clear()
        config = self.settings.get("glpi_api", {})

        if not config.get("enabled"):
            print(
                "GLPI API desabilitada. Configure somente "
                "em settings.local.json."
            )
            self.pause()
            return

        ticket = input("ID do chamado GLPI: ").strip()
        if not ticket.isdigit():
            return

        valid_report = (
            self.context.belongs_to(self.host)
            and self.context.report_path is not None
            and self.context.report_path.exists()
        )
        if not valid_report:
            print(
                "Gere um relatório para a estação selecionada "
                "antes de enviar ao GLPI."
            )
            self.pause()
            return

        client = GLPIClient(
            config.get("base_url", ""),
            config.get("app_token", ""),
            config.get("user_token", ""),
        )
        logger = getattr(self.executor, "logger", None)

        try:
            client.add_ticket_followup(
                int(ticket),
                self.context.report_path.read_text(
                    encoding="utf-8"
                ),
            )
            if logger:
                with logger.bind(
                    correlation_id=self.context.correlation_id
                ):
                    logger.log_event(
                        "glpi_ticket_followup",
                        self.host,
                        "success",
                        ticket_id=int(ticket),
                    )
            print("✓ Acompanhamento enviado ao GLPI.")
        except GLPIError as exc:
            if logger:
                with logger.bind(
                    correlation_id=self.context.correlation_id
                ):
                    logger.log_event(
                        "glpi_ticket_followup",
                        self.host,
                        "failure",
                        ticket_id=int(ticket),
                        error=str(exc),
                    )
            print(f"✗ GLPI: {exc}")
        finally:
            try:
                client.kill_session()
            except Exception:
                pass

        self.pause()

    @staticmethod
    def _probe_payload(result: CommandResult) -> dict[str, Any]:
        if result.success and isinstance(result.data, dict):
            return {
                **result.data,
                "_probe_success": True,
                "_transport": result.transport,
            }

        return {
            "_probe_success": False,
            "_transport": result.transport,
            "_error": result.stderr,
            "_indeterminate": result.indeterminate,
        }

    def _health_probe(self, host: str) -> dict[str, Any]:
        return self._probe_payload(self.health.snapshot(host))

    def _cleanup_probe(self, host: str) -> dict[str, Any]:
        return self._probe_payload(
            self.disk.cleanup_estimate(host)
        )

    def _spooler_probe(self, host: str) -> dict[str, Any]:
        return self._probe_payload(
            self.printers.spooler_status(host)
        )

    def _gpresult_probe(self, host: str) -> dict[str, Any]:
        result = self.domain.gpresult(host)
        return {
            "_probe_success": result.success,
            "_transport": result.transport,
            "_stdout": result.stdout,
            "_error": result.stderr,
            "_indeterminate": result.indeterminate,
        }

    def menu_remediations(self) -> None:
        if not self.require_host():
            return

        self.clear()
        print("REMEDIAÇÕES GUIADAS")
        print("1 - Limpeza segura de temporários")
        print("2 - Reiniciar Spooler")
        print("3 - Resetar componentes do Windows Update")
        print("4 - GPUpdate /force")
        print("0 - Voltar")

        option = input("Opção: ").strip()

        specs = {
            "1": {
                "spec": RemediationSpec(
                    "safe_cleanup",
                    "Limpeza segura de temporários",
                    "médio",
                    True,
                    False,
                    False,
                    (
                        "Não há rollback automático para "
                        "temporários removidos. A lixeira não é tocada."
                    ),
                ),
                "action": self.disk.cleanup_safe,
                "operation_class": OperationClass.HEAVY_WRITE,
                "before_probe": self._cleanup_probe,
                "after_probe": self._cleanup_probe,
                "validator": validate_cleanup,
            },
            "2": {
                "spec": RemediationSpec(
                    "restart_spooler",
                    "Reiniciar Spooler",
                    "baixo",
                    True,
                ),
                "action": self.printers.restart_spooler,
                "operation_class": OperationClass.LIGHT_WRITE,
                "before_probe": self._spooler_probe,
                "after_probe": self._spooler_probe,
                "validator": validate_spooler,
            },
            "3": {
                "spec": RemediationSpec(
                    "reset_windows_update",
                    "Resetar componentes do Windows Update",
                    "alto",
                    True,
                    False,
                    False,
                    (
                        "SoftwareDistribution e catroot2 são "
                        "renomeados com timestamp; serviços "
                        "originalmente ativos são restaurados."
                    ),
                ),
                "action": self.updates.reset_components,
                "operation_class": OperationClass.HEAVY_WRITE,
                "before_probe": self._health_probe,
                "after_probe": self._health_probe,
                "validator": validate_windows_update_reset,
            },
            "4": {
                "spec": RemediationSpec(
                    "gpupdate_force",
                    "GPUpdate /force",
                    "baixo",
                    True,
                ),
                "action": self.system.gpupdate,
                "operation_class": OperationClass.LIGHT_WRITE,
                "before_probe": self._gpresult_probe,
                "after_probe": self._gpresult_probe,
                "validator": validate_gpupdate,
            },
        }

        item = specs.get(option)
        if not item:
            return

        spec: RemediationSpec = item["spec"]
        print(f"\nAção: {spec.title}")
        print(
            f"Impacto: {spec.impact.upper()} | "
            f"Reboot esperado: "
            f"{'SIM' if spec.requires_reboot else 'NÃO'}"
        )
        if spec.rollback:
            print(f"Rollback/observação: {spec.rollback}")

        if spec.requires_confirmation and not self.confirm(
            f"Executar {spec.title} em {self.host}?"
        ):
            return

        try:
            remediation = self.jobs.run(
                f"Remediação: {spec.title}",
                self._trace(
                    lambda: self.remediation_engine.execute(
                        self.host,
                        spec,
                        item["action"],
                        before_probe=item["before_probe"],
                        after_probe=item["after_probe"],
                        validator=item["validator"],
                    ),
                    action=f"remediation.{spec.key}",
                ),
                operation_class=item["operation_class"],
                timeout=self.long_timeout,
                host=self.host,
                correlation_id=self.context.correlation_id,
            )
        except Exception as exc:
            print(
                f"✗ Remediação falhou: "
                f"{type(exc).__name__}: {exc}"
            )
            self.pause()
            return

        self.context.remediation = remediation

        self.show_result(remediation.command_result)
        print(
            f"\nVALIDAÇÃO: {remediation.validation.status.value} — "
            f"{remediation.validation.message}"
        )

        changes = (
            diff_values(remediation.before, remediation.after)
            if remediation.after is not None
            else []
        )
        print("\nANTES / DEPOIS")
        if changes:
            for change in changes[:50]:
                print(
                    f"- {change['path']}: {change['before']} -> "
                    f"{change['after']}"
                )
        else:
            print(
                "Nenhuma diferença relevante foi detectada "
                "pelas sondas desta remediação."
            )

        if self.db:
            validated = (
                remediation.validation.status
                is ValidationStatus.PASS
            )
            self.db.save_remediation(
                self.host,
                spec.key,
                validated,
                asdict(remediation),
                correlation_id=self.context.correlation_id,
            )
            if isinstance(remediation.after, dict):
                self.db.save_snapshot(
                    self.host,
                    remediation.after,
                    kind=f"after:{spec.key}",
                    correlation_id=self.context.correlation_id,
                )

        self.pause()

    @staticmethod
    def _execution_evidence(value: Any) -> Any:
        if isinstance(value, CommandResult):
            if value.data is not None:
                return value.data
            return {
                "success": value.success,
                "transport": value.transport,
                "stdout": value.stdout,
                "stderr": value.stderr,
                "indeterminate": value.indeterminate,
            }
        return value

    def _execution_policy_context(self) -> ExecutionPolicyContext:
        if not self.context.session:
            raise ExecutionBlockedError(
                "Sessão lógica não está disponível."
            )
        return ExecutionPolicyContext.from_session(
            self.context.session
        )

    def _recover_execution_host(
        self,
        host: str,
        timeout_seconds: int,
        delay_seconds: int,
    ) -> RecoveryResult:
        if delay_seconds > 0:
            time.sleep(delay_seconds)

        started = time.monotonic()
        attempts = 0
        last_error: str | None = None
        last_state: str | None = None

        while time.monotonic() - started < timeout_seconds:
            attempts += 1
            try:
                session = self.sessions.open(
                    host,
                    refresh=True,
                )
                last_state = str(
                    session.connectivity.get("state") or ""
                )
                if session.ready:
                    self.context.session = session
                    return RecoveryResult(
                        attempted=True,
                        ready=True,
                        attempts=attempts,
                        elapsed_seconds=round(
                            time.monotonic() - started,
                            2,
                        ),
                        transport=session.transport,
                        state=last_state,
                    )
            except Exception as exc:
                last_error = (
                    f"{type(exc).__name__}: {exc}"
                )

            time.sleep(5)

        return RecoveryResult(
            attempted=True,
            ready=False,
            attempts=attempts,
            elapsed_seconds=round(
                time.monotonic() - started,
                2,
            ),
            state=last_state,
            error=last_error or "Prazo de recuperação excedido.",
        )

    @staticmethod
    def _list_payload(result: CommandResult) -> list[dict[str, Any]]:
        data = result.data
        if isinstance(data, list):
            return [
                item
                for item in data
                if isinstance(item, dict)
            ]
        if isinstance(data, dict):
            return [data]
        return []

    def _selector_options(
        self,
        selector: SelectorKind,
        parameter_key: str,
    ) -> list[tuple[Any, str]]:
        if not self.host:
            return []

        def load(label: str, callback):
            try:
                return self.jobs.run(
                    label,
                    self._trace(
                        callback,
                        action=f"execution.selector.{selector.value.casefold()}",
                    ),
                    timeout=300,
                    operation_class=OperationClass.READ_ONLY,
                    host=self.host,
                    correlation_id=self.context.correlation_id,
                )
            except Exception:
                return None

        options: list[tuple[Any, str]] = []

        if selector is SelectorKind.PROCESS:
            result = load(
                "Carregando processos",
                lambda: self.system.list_processes(
                    self.host,
                    100,
                ),
            )
            if isinstance(result, CommandResult) and result.success:
                for item in self._list_payload(result):
                    name = item.get("Name")
                    pid = item.get("Id")
                    if name is None or pid is None:
                        continue
                    value = (
                        int(pid)
                        if parameter_key == "pid"
                        else str(name)
                    )
                    memory = item.get("WorkingSet")
                    detail = (
                        f"{name} | PID {pid}"
                        + (
                            f" | RAM {int(memory) // 1048576} MB"
                            if isinstance(memory, (int, float))
                            else ""
                        )
                    )
                    options.append((value, detail))

        elif selector is SelectorKind.SERVICE:
            result = load(
                "Carregando serviços",
                lambda: self.system.list_services(self.host),
            )
            if isinstance(result, CommandResult) and result.success:
                for item in self._list_payload(result):
                    name = item.get("Name")
                    if not name:
                        continue
                    options.append(
                        (
                            str(name),
                            (
                                f"{item.get('DisplayName') or name} | "
                                f"{name} | {item.get('Status')} | "
                                f"{item.get('StartType')}"
                            ),
                        )
                    )

        elif selector is SelectorKind.ADAPTER:
            result = load(
                "Carregando adaptadores",
                lambda: self.network.adapters(self.host),
            )
            if isinstance(result, CommandResult) and result.success:
                for item in self._list_payload(result):
                    name = item.get("Name")
                    if name:
                        options.append(
                            (
                                str(name),
                                (
                                    f"{name} | {item.get('Status')} | "
                                    f"{item.get('LinkSpeed') or '-'}"
                                ),
                            )
                        )

        elif selector is SelectorKind.PRINTER:
            result = load(
                "Carregando impressoras",
                lambda: self.printers.list(self.host),
            )
            if isinstance(result, CommandResult) and result.success:
                for item in self._list_payload(result):
                    name = item.get("Name")
                    if name:
                        options.append(
                            (
                                str(name),
                                (
                                    f"{name} | "
                                    f"{item.get('PrinterStatus') or '-'} | "
                                    f"{item.get('DriverName') or '-'}"
                                ),
                            )
                        )

        elif selector is SelectorKind.PROFILE:
            result = load(
                "Carregando perfis",
                lambda: self.users.profiles(self.host),
            )
            if isinstance(result, CommandResult) and result.success:
                for item in self._list_payload(result):
                    sid = item.get("SID")
                    if sid:
                        options.append(
                            (
                                str(sid),
                                (
                                    f"{item.get('LocalPath') or sid} | "
                                    f"{sid} | Loaded={item.get('Loaded')}"
                                ),
                            )
                        )

        elif selector is SelectorKind.DEVICE:
            result = load(
                "Carregando dispositivos PnP",
                lambda: self.devices.present_devices(self.host),
            )
            if isinstance(result, CommandResult) and result.success:
                for item in self._list_payload(result):
                    instance_id = item.get("InstanceId")
                    if instance_id:
                        options.append(
                            (
                                str(instance_id),
                                (
                                    f"{item.get('FriendlyName') or instance_id} | "
                                    f"{item.get('Class') or '-'} | "
                                    f"{item.get('Status') or '-'}"
                                ),
                            )
                        )

        elif selector is SelectorKind.SESSION:
            result = load(
                "Carregando sessões",
                lambda: self.system.sessions(self.host),
            )
            if isinstance(result, CommandResult) and result.success:
                for line in result.stdout.splitlines()[1:]:
                    values = re.findall(r"\b\d+\b", line)
                    if not values:
                        continue
                    session_id = int(values[0])
                    options.append(
                        (
                            session_id,
                            re.sub(r"\s+", " ", line.strip()),
                        )
                    )

        return options[:150]

    def _prompt_execution_parameters(self, bound) -> dict[str, Any]:
        raw: dict[str, Any] = {}
        for parameter in bound.spec.parameters:
            if parameter.help_text:
                print(f"  ℹ {parameter.help_text}")

            default_text = (
                f" [padrão: {parameter.default}]"
                if parameter.default is not None
                else ""
            )

            if parameter.selector is not None:
                options = self._selector_options(
                    parameter.selector,
                    parameter.key,
                )
                if options:
                    print(f"\n{parameter.label}:")
                    for index, (_, label) in enumerate(options, 1):
                        print(f"  {index:3} - {label}")
                    print("  M - Informar manualmente")
                    if not parameter.required:
                        print("  0 - Nenhum / todos quando aplicável")

                    selected = input("Escolha: ").strip()
                    if (
                        not parameter.required
                        and selected == "0"
                    ):
                        raw[parameter.key] = None
                        continue
                    if (
                        selected.isdigit()
                        and 1 <= int(selected) <= len(options)
                    ):
                        raw[parameter.key] = options[
                            int(selected) - 1
                        ][0]
                        continue
                    if selected.casefold() != "m":
                        raw[parameter.key] = selected
                        continue

            if parameter.kind is ParameterKind.CHOICE:
                print(f"\n{parameter.label}:")
                for index, option in enumerate(parameter.choices, 1):
                    print(f"  {index} - {option}")
                choice = input(
                    f"Escolha{default_text}: "
                ).strip()
                if not choice and parameter.default is not None:
                    raw[parameter.key] = parameter.default
                    continue
                if (
                    choice.isdigit()
                    and 1 <= int(choice) <= len(parameter.choices)
                ):
                    raw[parameter.key] = parameter.choices[
                        int(choice) - 1
                    ]
                else:
                    raw[parameter.key] = choice
                continue

            suffix = "" if parameter.required else " [opcional]"
            value = input(
                f"{parameter.label}{suffix}{default_text}: "
            )
            raw[parameter.key] = value

        return raw

    def _show_execution_plan(self, plan) -> None:
        print("\nPRECONDITIONS / CAPABILITIES")
        symbols = {
            PolicyState.PASS: "✓",
            PolicyState.FAIL: "✗",
            PolicyState.WARN: "⚠",
        }
        for item in plan.policy.checks:
            print(
                f" {symbols[item.state]} "
                f"[{item.state.value}] {item.message}"
            )
        print(
            "\nPlano: "
            f"{'LIBERADO' if plan.allowed else 'BLOQUEADO'}"
        )

    def _confirm_execution(self, bound) -> bool:
        spec = bound.spec
        strong = (
            spec.risk in {RiskLevel.HIGH, RiskLevel.CRITICAL}
            or spec.destructive
            or spec.may_break_connectivity
        )

        if not spec.requires_confirmation:
            return True

        if not strong:
            return self.confirm(
                f"Executar {spec.title} em {self.host}?"
            )

        expected = f"EXECUTAR {self.host}"
        print("\n⚠ CONFIRMAÇÃO REFORÇADA")
        print(
            f"Risco: {spec.risk.value} | "
            f"Destrutiva: {'SIM' if spec.destructive else 'NÃO'} | "
            f"Desconexão: {spec.disconnect_mode.value}"
        )
        print(f"Impacto: {spec.impact}")
        typed = input(
            f"Digite exatamente '{expected}' para confirmar: "
        ).strip()
        return typed.casefold() == expected.casefold()

    def _show_execution_action(self, bound) -> None:
        spec = bound.spec
        flags: list[str] = []
        if spec.destructive:
            flags.append("DESTRUTIVA")
        if spec.requires_reboot:
            flags.append("REBOOT")
        if spec.may_break_connectivity:
            flags.append("REDE")
        if spec.rollback_strategy:
            flags.append("ROLLBACK")
        flag_text = f" | {', '.join(flags)}" if flags else ""

        print(f"\n=== {spec.title} ===")
        print(
            f"Categoria: {spec.category_label} | "
            f"Risco: {spec.risk.value} | "
            f"Classe: {spec.operation_class.value}{flag_text}"
        )
        print(f"Descrição: {spec.description}")
        print(f"Impacto: {spec.impact}")
        print(f"Timeout: {spec.timeout_seconds}s")
        print(
            f"Transportes: {', '.join(spec.allowed_transports)} | "
            f"Idempotente: {'SIM' if spec.idempotent else 'NÃO'} | "
            f"Retry: {spec.retry_policy.value}"
        )
        if spec.required_capabilities:
            print(
                "Capabilities: "
                + ", ".join(spec.required_capabilities)
            )
        if spec.rollback_strategy:
            print(f"Rollback: {spec.rollback_strategy}")
        if spec.recommendation:
            print(f"Recomendação: {spec.recommendation}")

    def _choose_execution_action(
        self,
        actions,
        *,
        title: str,
    ):
        if not actions:
            print("Nenhuma ação encontrada.")
            self.pause()
            return None

        self.clear()
        print(title)
        print(f"Alvo: {self.host}\n")
        for index, bound in enumerate(actions, 1):
            spec = bound.spec
            flags = []
            if spec.destructive:
                flags.append("D")
            if spec.requires_reboot:
                flags.append("R")
            if spec.may_break_connectivity:
                flags.append("C")
            if spec.rollback_strategy:
                flags.append("U")
            suffix = f" [{' '.join(flags)}]" if flags else ""
            print(
                f"{index:2} - [{spec.risk.value:<8}] "
                f"{spec.title}{suffix}"
            )

        print(
            "\nD=destrutiva | R=reboot | C=conectividade | "
            "U=rollback disponível"
        )
        print("0 - Voltar")
        option = input("Ação: ").strip()
        if (
            option == "0"
            or not option.isdigit()
            or not 1 <= int(option) <= len(actions)
        ):
            return None
        return actions[int(option) - 1]

    def _rollback_last_execution(self) -> None:
        record = (
            self.context.rollback_stack[-1]
            if self.context.rollback_stack
            else None
        )
        if (
            record is None
            or not record.rollback_available
            or record.rollback_result is not None
        ):
            print("Não há execução reversível pendente neste atendimento.")
            self.pause()
            return

        expected = f"DESFAZER {self.host}"
        print(f"\nRollback: {record.action.title}")
        print(f"Estratégia: {record.action.rollback_strategy}")
        typed = input(
            f"Digite exatamente '{expected}' para confirmar: "
        ).strip()
        if typed.casefold() != expected.casefold():
            print("Rollback cancelado.")
            self.pause()
            return

        try:
            rollback = self.jobs.run(
                f"Rollback: {record.action.title}",
                self._trace(
                    lambda: self.execution_engine.rollback(
                        record,
                        context=self._execution_policy_context(),
                        operator=getpass.getuser(),
                    ),
                    action=f"execution.rollback.{record.action.key}",
                ),
                operation_class=record.action.operation_class,
                timeout=record.action.timeout_seconds,
                host=self.host,
                correlation_id=self.context.correlation_id,
            )
        except Exception as exc:
            print(
                f"✗ Rollback falhou: "
                f"{type(exc).__name__}: {exc}"
            )
            self.pause()
            return

        self.show_result(rollback.result)
        print(
            f"\nVALIDAÇÃO DO ROLLBACK: "
            f"{rollback.validation.status.value} — "
            f"{rollback.validation.message}"
        )

        if self.db and record.execution_id is not None:
            self.db.save_rollback_record(
                rollback,
                original_execution_id=record.execution_id,
                operator=getpass.getuser(),
                correlation_id=self.context.correlation_id,
            )

        if rollback.validation.status is ValidationStatus.PASS:
            record.rollback_available = False
            if (
                self.context.rollback_stack
                and self.context.rollback_stack[-1] is record
            ):
                self.context.rollback_stack.pop()

        self.pause()

    def _execute_bound_action(self, bound) -> None:
        self._show_execution_action(bound)

        try:
            raw_parameters = self._prompt_execution_parameters(
                bound
            )
            parsed = self.execution_engine.validate_parameters(
                bound,
                raw_parameters,
            )
            plan = self.execution_engine.plan(
                self.host,
                bound.spec.key,
                parsed,
                context=self._execution_policy_context(),
            )
        except (ValueError, ExecutionBlockedError) as exc:
            print(f"\n✗ Plano inválido: {exc}")
            self.pause()
            return

        if parsed:
            print("\nParâmetros:")
            sensitive = {
                item.key
                for item in bound.spec.parameters
                if item.sensitive
            }
            for key, value in parsed.items():
                visible = "***" if key in sensitive else value
                print(f" - {key}: {visible}")

        self._show_execution_plan(plan)
        if not plan.allowed:
            print(
                "\nAção bloqueada antes da confirmação. "
                "Corrija capabilities/preconditions e tente novamente."
            )
            self.pause()
            return

        if not self._confirm_execution(bound):
            print("Execução cancelada.")
            self.pause()
            return

        try:
            record = self.jobs.run(
                f"Execução: {bound.spec.title}",
                self._trace(
                    lambda: self.execution_engine.execute(
                        self.host,
                        bound.spec.key,
                        parsed,
                        context=self._execution_policy_context(),
                        operator=getpass.getuser(),
                        reconnect=self._recover_execution_host,
                    ),
                    action=f"execution.{bound.spec.key}",
                ),
                operation_class=bound.spec.operation_class,
                timeout=bound.spec.timeout_seconds,
                host=self.host,
                correlation_id=self.context.correlation_id,
            )
        except Exception as exc:
            print(
                f"\n✗ Execução falhou: "
                f"{type(exc).__name__}: {exc}"
            )
            self.pause()
            return

        remediation = record.remediation
        self.context.remediation = remediation
        self.context.execution = record
        if record.rollback_available:
            self.context.rollback_stack.append(record)
        self.show_result(remediation.command_result)

        if record.recovery is not None:
            print(
                f"\nRECUPERAÇÃO: "
                f"{'READY' if record.recovery.ready else 'NÃO RECUPERADO'} | "
                f"tentativas={record.recovery.attempts} | "
                f"{record.recovery.elapsed_seconds:.1f}s | "
                f"{record.recovery.transport or '-'}"
            )

        print(
            f"\nVALIDAÇÃO: "
            f"{remediation.validation.status.value} — "
            f"{remediation.validation.message}"
        )

        before = self._execution_evidence(remediation.before)
        after = self._execution_evidence(remediation.after)
        print("\nANTES:")
        print(
            json.dumps(
                before,
                indent=2,
                ensure_ascii=False,
                default=str,
            )
            if before is not None
            else "-"
        )
        print("\nDEPOIS:")
        print(
            json.dumps(
                after,
                indent=2,
                ensure_ascii=False,
                default=str,
            )
            if after is not None
            else "-"
        )

        if self.db:
            self.db.save_execution_record(
                record,
                correlation_id=self.context.correlation_id,
            )
            if isinstance(after, dict):
                self.db.save_snapshot(
                    self.host,
                    after,
                    kind=f"execution:{bound.spec.key}",
                    correlation_id=self.context.correlation_id,
                )

        if record.rollback_available:
            print(
                "\n↩ Rollback disponível. Volte à Central de "
                "Execuções e escolha [U] para desfazer."
            )

        print(
            f"\nCorrelation ID: "
            f"{self.context.correlation_id}"
        )
        self.pause()

    def open_execution_category(self, category: str) -> None:
        self.menu_execution(category=category)

    def menu_execution(self, category: str | None = None) -> None:
        if not self.require_host():
            return
        if not self.context.session or not self.context.session.ready:
            print(
                "Nenhum transporte administrativo validado para a estação. "
                "Execute o preflight/conectividade antes de usar a Central de Execuções."
            )
            self.pause()
            return

        validation = self.settings.get(
            "_execution_catalog_validation",
            {},
        )

        if category is None:
            self.clear()
            print("CENTRAL DE EXECUÇÕES")
            print(
                f"Alvo: {self.host} | "
                f"Transporte: {self.context.session.transport} | "
                f"Ações: {len(self.execution_registry)}"
            )
            counts = validation.get("enabled_counts", {})
            if counts:
                print(
                    "Catálogos: "
                    f"pacotes={counts.get('packages', 0)} | "
                    f"certificados={counts.get('certificates', 0)} | "
                    f"registro={counts.get('registry_actions', 0)}"
                )
            issues = validation.get("issues", [])
            if issues:
                print(
                    f"⚠ {len(issues)} entrada(s) de catálogo "
                    "foram desabilitadas no bootstrap."
                )

            categories = self.execution_registry.categories()
            for index, (key, label) in enumerate(categories, 1):
                count = len(
                    self.execution_registry.by_category(key)
                )
                print(f"{index} - {label} ({count})")

            print("B - Buscar ação")
            if (
                bool(self.context.rollback_stack)
                and self.context.rollback_stack[-1].rollback_available
                and self.context.rollback_stack[-1].rollback_result is None
            ):
                print("U - Desfazer última execução reversível")
            print("0 - Voltar")

            option = input("Categoria/Ação: ").strip()
            if option.casefold() == "b":
                query = input("Buscar: ").strip()
                bound = self._choose_execution_action(
                    self.execution_registry.search(query),
                    title=f"BUSCA — {query or 'todas as ações'}",
                )
                if bound:
                    self._execute_bound_action(bound)
                return

            if option.casefold() == "u":
                self._rollback_last_execution()
                return

            if (
                option == "0"
                or not option.isdigit()
                or not 1 <= int(option) <= len(categories)
            ):
                return
            category = categories[int(option) - 1][0]

        actions = self.execution_registry.by_category(category)
        label = (
            actions[0].spec.category_label
            if actions
            else category
        )
        bound = self._choose_execution_action(
            actions,
            title=f"EXECUÇÕES — {label}",
        )
        if bound:
            self._execute_bound_action(bound)

    def menu_baseline(self) -> None:
        self.clear()
        profiles = self.baselines.available()
        print(f"Perfil ativo: {self.active_baseline_profile}")

        for index, profile in enumerate(profiles, 1):
            print(f"{index} - {profile}")
        print("0 - Voltar")

        choice = input("Perfil: ").strip()
        if (
            choice == "0"
            or not choice.isdigit()
            or not 1 <= int(choice) <= len(profiles)
        ):
            return

        self.active_baseline_profile = profiles[int(choice) - 1]
        print(
            "✓ Baseline ativo nesta sessão: "
            f"{self.active_baseline_profile}"
        )
        self.pause()
