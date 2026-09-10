from __future__ import annotations

import json
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
        self.report_builder = SupportReportBuilder()
        self.report_exporter = ReportExporter(root / "reports" / "support")

        update_config = settings.get("updates", {})
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
                f"UNKNOWN {compliance['unknown']}"
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
        if self.context.remediation:
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
