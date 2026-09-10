from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any, Callable

from core.actions import ActionSpec
from core.config import ConfigLoader
from core.jobs import JobManager, OperationClass, ResponsiveJobRunner
from core.result import CommandResult
from core.validation import validate_host, validate_process_name, validate_windows_path
from modules.crashes import CrashesModule
from modules.devices import DevicesModule
from modules.diagnostic_package import DiagnosticPackageModule
from modules.diagnostics import DiagnosticsModule
from modules.disk import DiskModule
from modules.domain import DomainModule
from modules.glpi import GLPIModule
from modules.health import HealthModule, calculate_health_score
from modules.network import NetworkModule
from modules.performance import PerformanceModule
from modules.printers import PrintersModule
from modules.repair import RepairModule
from modules.security import SecurityModule
from modules.software import SoftwareModule
from modules.startup import StartupModule
from modules.storage import StorageModule
from modules.sysinternals import SysinternalsModule
from modules.system import SystemModule
from modules.tasks import TasksModule
from modules.updates import UpdatesModule
from modules.users_profiles import UsersProfilesModule
from modules.workstation_tools import WorkstationToolsModule


class ConsoleBase:
    """Base operacional limpa da UI v5.

    A geração v3 permanece no repositório como compatibilidade histórica; a v5
    usa esta classe para evitar herdar decisões antigas de configuração/scheduler.
    """

    def __init__(
        self,
        executor,
        settings_path: Path,
        *,
        settings: dict[str, Any] | None = None,
        job_manager: JobManager | None = None,
    ) -> None:
        self.executor = executor
        self.settings_path = settings_path
        self.settings = settings or ConfigLoader(settings_path).settings

        ui = self.settings.get("ui", {})
        self.long_timeout = int(
            ui.get("long_operation_timeout_seconds", 3600)
        )
        self.jobs = ResponsiveJobRunner(
            heartbeat_seconds=float(ui.get("heartbeat_seconds", 0.2)),
            manager=job_manager,
        )

        self.host: str | None = None
        self.health_snapshot: dict[str, Any] | None = None

        self.diag = DiagnosticsModule(executor)
        self.health = HealthModule(executor)
        self.network = NetworkModule(executor)
        self.system = SystemModule(executor)
        self.software = SoftwareModule(
            executor,
            settings_path,
            settings=self.settings,
        )
        self.glpi = GLPIModule(
            executor,
            settings_path,
            settings=self.settings,
        )
        self.security = SecurityModule(executor)
        self.updates = UpdatesModule(executor)
        self.users = UsersProfilesModule(executor)
        self.printers = PrintersModule(executor)
        self.domain = DomainModule(executor)
        self.disk = DiskModule(executor)
        self.packager = DiagnosticPackageModule(
            executor,
            settings_path.parent.parent / "reports" / "diagnostics",
        )
        self.repair = RepairModule(executor)
        self.devices = DevicesModule(executor)
        self.performance = PerformanceModule(executor)
        self.startup = StartupModule(executor)
        self.crashes = CrashesModule(executor)
        self.tasks = TasksModule(executor)
        self.storage = StorageModule(executor)
        self.sysinternals = SysinternalsModule(
            executor,
            self.settings.get("sysinternals_dir", r"C:\Sysinternals"),
        )
        self.tools = WorkstationToolsModule(executor)

    @staticmethod
    def clear() -> None:
        os.system("cls" if os.name == "nt" else "clear")

    @staticmethod
    def pause() -> None:
        input("\nPressione ENTER para continuar...")

    @staticmethod
    def confirm(text: str) -> bool:
        return input(f"\n⚠ {text} [digite SIM]: ").strip().upper() == "SIM"

    def require_host(self) -> bool:
        if self.host:
            return True
        print("Selecione uma estação primeiro.")
        self.pause()
        return False

    @staticmethod
    def _display_value(value: Any) -> str:
        if value is None or value == "":
            return "-"
        if isinstance(value, bool):
            return "SIM" if value else "NÃO"
        if isinstance(value, (dict, list)):
            return json.dumps(value, ensure_ascii=False, default=str)
        return str(value)

    @staticmethod
    def _clip(text: str, width: int) -> str:
        if len(text) <= width:
            return text
        if width <= 1:
            return text[:width]
        return text[: width - 1] + "…"

    @classmethod
    def _print_table(
        cls,
        rows: list[dict[str, Any]],
        columns: list[tuple[str, str, int]] | None = None,
    ) -> None:
        if not rows:
            print("Nenhum registro encontrado.")
            return

        if columns is None:
            keys: list[str] = []
            for row in rows:
                for key in row:
                    if key not in keys:
                        keys.append(key)
            keys = keys[:8]
            columns = []
            for key in keys:
                longest = max(
                    [len(str(key))]
                    + [
                        len(cls._display_value(row.get(key)))
                        for row in rows[:100]
                    ]
                )
                columns.append(
                    (key, key, max(8, min(longest, 28)))
                )

        header = " | ".join(
            cls._clip(label, width).ljust(width)
            for _, label, width in columns
        )
        divider = "-+-".join("-" * width for _, _, width in columns)
        print(header)
        print(divider)

        for row in rows:
            print(
                " | ".join(
                    cls._clip(
                        cls._display_value(row.get(key)),
                        width,
                    ).ljust(width)
                    for key, _, width in columns
                )
            )

    @classmethod
    def _show_drivers(cls, data: Any) -> None:
        rows = data if isinstance(data, list) else [data]
        rows = [row for row in rows if isinstance(row, dict)]
        columns = [
            ("DeviceName", "Dispositivo", 36),
            ("Manufacturer", "Fabricante", 27),
            ("DriverVersion", "Versão", 18),
            ("DriverDate", "Data", 10),
            ("IsSigned", "Assinado", 8),
            ("InfName", "INF", 14),
            ("Count", "Qtd.", 4),
        ]
        print("\nDRIVERS INSTALADOS")
        cls._print_table(rows, columns)
        total = sum(int(row.get("Count") or 1) for row in rows)
        unsigned = sum(
            int(row.get("Count") or 1)
            for row in rows
            if row.get("IsSigned") is False
        )
        print(
            f"\nResumo: {total} instâncias | "
            f"{len(rows)} entradas agrupadas | "
            f"{unsigned} não assinada(s)"
        )

    @classmethod
    def show_result(cls, result: CommandResult) -> None:
        status = "✓ SUCESSO" if result.success else "✗ FALHA"
        print(
            f"\n{status} [{result.transport}] — "
            f"{result.duration_ms} ms"
        )

        if result.data is not None:
            if result.metadata.get("view") == "drivers":
                cls._show_drivers(result.data)
            elif (
                isinstance(result.data, list)
                and all(
                    isinstance(item, dict)
                    for item in result.data
                )
            ):
                print()
                cls._print_table(result.data)
            else:
                print(
                    json.dumps(
                        result.data,
                        indent=2,
                        ensure_ascii=False,
                        default=str,
                    )
                )
        elif result.stdout:
            print(result.stdout)

        if result.stderr:
            label = "Erro" if not result.success else "Aviso"
            print(f"\n{label}: {result.stderr}")

        if result.indeterminate:
            reason = result.metadata.get(
                "indeterminate_reason",
                "resultado remoto não confirmado",
            )
            print(f"\n⚠ RESULTADO INDETERMINADO: {reason}")
            print(
                "Valide o estado da estação antes de repetir a ação."
            )

    @staticmethod
    def _legacy_key(title: str) -> str:
        value = re.sub(r"[^a-z0-9]+", ".", title.casefold()).strip(".")
        return f"ui.{value or 'action'}"

    def execute(
        self,
        action: str | ActionSpec,
        func: Callable[[], CommandResult],
        *,
        timeout: int | None = None,
        operation_class: OperationClass | None = None,
    ) -> CommandResult | None:
        if isinstance(action, ActionSpec):
            spec = action
        else:
            spec = ActionSpec(
                key=self._legacy_key(action),
                title=action,
                operation_class=(
                    operation_class or OperationClass.READ_ONLY
                ),
                timeout_seconds=timeout,
            )

        effective_timeout = (
            timeout
            or spec.timeout_seconds
            or self.long_timeout
        )
        print(f"\n▶ {spec.title}")

        runner = func
        trace = getattr(self, "_trace", None)
        if callable(trace):
            runner = trace(func, action=spec.key)

        try:
            result = self.jobs.run(
                spec.title,
                runner,
                timeout=effective_timeout,
                operation_class=spec.operation_class,
                host=self.host or "local-ui",
                correlation_id=getattr(
                    self,
                    "correlation_id",
                    None,
                ),
            )
        except TimeoutError as exc:
            print(f"\n✗ TIMEOUT: {exc}")
            return None
        except Exception as exc:
            print(
                f"\n✗ ERRO INTERNO: {type(exc).__name__}: {exc}"
            )
            return None

        if not isinstance(result, CommandResult):
            print(
                "\n✗ ERRO DE CONTRATO: a operação não retornou CommandResult."
            )
            return None

        self.show_result(result)
        return result

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

        self.host = host
        result = self.execute(
            ActionSpec.read(
                "health.preflight",
                "Pré-flight da estação",
                timeout_seconds=120,
            ),
            lambda: self.health.snapshot(host),
        )
        if (
            result
            and result.success
            and isinstance(result.data, dict)
        ):
            self.health_snapshot = result.data
            self.show_health(result.data)
        self.pause()

    def show_health(self, snapshot: dict[str, Any]) -> None:
        baseline = self.settings.get("compliance", {})
        health = calculate_health_score(
            snapshot,
            baseline=baseline,
        )
        print("\n=== SAÚDE DA ESTAÇÃO ===")
        print(
            f"Score: {health['score']}/100 | "
            f"Estado: {health['overall_state']} | "
            f"CPU: {snapshot.get('CPUPercent','-')}% | "
            f"RAM: {snapshot.get('RAMUsedPercent','-')}% | "
            f"Disco livre: {snapshot.get('DiskFreePercent','-')}% | "
            f"Uptime: {snapshot.get('UptimeDays','-')} dias"
        )

    def menu_performance(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "PERFORMANCE\n"
            "1 - Amostragem rápida (8s)\n"
            "2 - Amostragem detalhada (20s)\n"
            "0 - Voltar"
        )
        option = input("Opção: ").strip()
        if option == "1":
            self.execute(
                ActionSpec.read(
                    "performance.quick",
                    "Amostrando CPU/RAM/disco/rede",
                    timeout_seconds=120,
                ),
                lambda: self.performance.snapshot(self.host, 8, 1),
            )
        elif option == "2":
            self.execute(
                ActionSpec.read(
                    "performance.detailed",
                    "Amostrando performance detalhada",
                    timeout_seconds=180,
                ),
                lambda: self.performance.snapshot(self.host, 20, 1),
            )
        else:
            return
        self.pause()

    def menu_repair(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "REPARO DO WINDOWS\n"
            "1 - SFC\n2 - DISM CheckHealth\n3 - DISM ScanHealth\n"
            "4 - DISM RestoreHealth\n5 - Analisar Component Store\n"
            "6 - Limpar Component Store\n7 - CHKDSK online\n"
            "8 - Verificar WMI\n0 - Voltar"
        )
        option = input("Opção: ").strip()
        actions = {
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
            "4": (
                ActionSpec(
                    "repair.dism.restore",
                    "DISM RestoreHealth",
                    OperationClass.HEAVY_WRITE,
                    3600,
                    True,
                ),
                self.repair.dism_restorehealth,
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

    def menu_devices(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "DISPOSITIVOS / DRIVERS\n"
            "1 - Dispositivos com erro\n2 - Drivers\n"
            "3 - USB presentes\n4 - Reexaminar dispositivos\n"
            "5 - Exportar drivers\n0 - Voltar"
        )
        option = input("Opção: ").strip()
        actions = {
            "1": (
                ActionSpec.read(
                    "devices.problems",
                    "Analisando dispositivos com erro",
                    timeout_seconds=120,
                ),
                self.devices.problem_devices,
            ),
            "2": (
                ActionSpec.read(
                    "devices.drivers",
                    "Inventariando drivers",
                    timeout_seconds=180,
                ),
                self.devices.drivers,
            ),
            "3": (
                ActionSpec.read(
                    "devices.usb",
                    "Inventariando USB",
                    timeout_seconds=120,
                ),
                self.devices.usb_devices,
            ),
            "4": (
                ActionSpec(
                    "devices.rescan",
                    "Reexaminando hardware",
                    OperationClass.LIGHT_WRITE,
                    180,
                ),
                self.devices.rescan,
            ),
            "5": (
                ActionSpec(
                    "devices.export_drivers",
                    "Exportando drivers",
                    OperationClass.HEAVY_WRITE,
                    1800,
                ),
                self.devices.export_drivers,
            ),
        }
        item = actions.get(option)
        if not item:
            return
        spec, function = item
        self.execute(spec, lambda: function(self.host))
        self.pause()

    def menu_startup_tasks(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "INICIALIZAÇÃO / TAREFAS\n"
            "1 - Visão de inicialização\n"
            "2 - Tarefas agendadas\n"
            "3 - Tarefas com falha\n0 - Voltar"
        )
        option = input("Opção: ").strip()
        actions = {
            "1": ("Analisando inicialização", self.startup.overview),
            "2": ("Listando tarefas", self.tasks.list_tasks),
            "3": ("Localizando tarefas com falha", self.tasks.failed_tasks),
        }
        item = actions.get(option)
        if not item:
            return
        title, function = item
        self.execute(title, lambda: function(self.host), timeout=240)
        self.pause()

    def menu_crashes(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "CRASHES / BSOD\n"
            "1 - BSOD e dumps\n"
            "2 - Crashes de aplicações\n"
            "0 - Voltar"
        )
        option = input("Opção: ").strip()
        if option == "1":
            self.execute(
                "Coletando histórico de BSOD",
                lambda: self.crashes.bsod_history(self.host),
                timeout=240,
            )
        elif option == "2":
            self.execute(
                "Coletando crashes de aplicações",
                lambda: self.crashes.app_crashes(self.host),
                timeout=240,
            )
        else:
            return
        self.pause()

    def menu_security(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "SEGURANÇA\n"
            "1 - Postura de segurança\n"
            "2 - Ameaças recentes\n0 - Voltar"
        )
        option = input("Opção: ").strip()
        if option == "1":
            self.execute(
                "Coletando postura de segurança",
                lambda: self.security.posture(self.host),
                timeout=180,
            )
        elif option == "2":
            self.execute(
                "Consultando ameaças",
                lambda: self.security.threats(self.host),
                timeout=180,
            )
        else:
            return
        self.pause()

    def menu_network(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "REDE\n1 - Interfaces\n2 - IP/Gateway/DNS\n"
            "3 - ARP\n4 - Conexões TCP\n5 - Flush DNS\n"
            "6 - Renovar DHCP\n0 - Voltar"
        )
        option = input("Opção: ").strip()
        actions = {
            "1": (
                ActionSpec.read("network.adapters", "Interfaces", timeout_seconds=180),
                self.network.adapters,
            ),
            "2": (
                ActionSpec.read("network.ip", "Configuração IP", timeout_seconds=180),
                self.network.ip_configuration,
            ),
            "3": (
                ActionSpec.read("network.arp", "Tabela ARP", timeout_seconds=180),
                self.network.arp_table,
            ),
            "4": (
                ActionSpec.read("network.connections", "Conexões TCP", timeout_seconds=180),
                self.network.connections,
            ),
            "5": (
                ActionSpec(
                    "network.flush_dns",
                    "Flush DNS",
                    OperationClass.LIGHT_WRITE,
                    180,
                ),
                self.network.flush_dns,
            ),
            "6": (
                ActionSpec(
                    "network.renew_dhcp",
                    "Renovar DHCP",
                    OperationClass.DISRUPTIVE,
                    180,
                    True,
                ),
                self.network.renew_dhcp,
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

    def menu_users(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "USUÁRIOS / PERFIS\n1 - Sessões\n"
            "2 - Administradores locais\n3 - Perfis e tamanho\n"
            "0 - Voltar"
        )
        option = input("Opção: ").strip()
        actions = {
            "1": (
                ActionSpec.read("users.sessions", "Sessões", timeout_seconds=180),
                self.system.sessions,
            ),
            "2": (
                ActionSpec.read("users.admins", "Administradores locais", timeout_seconds=180),
                self.users.local_admins,
            ),
            "3": (
                ActionSpec(
                    "users.profiles",
                    "Perfis",
                    OperationClass.HEAVY_READ,
                    600,
                ),
                self.users.profiles,
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
        self.clear()
        print(
            "SOFTWARE / GLPI\n1 - Software instalado\n"
            "2 - Winget\n3 - GLPI status\n"
            "4 - Forçar inventário GLPI\n0 - Voltar"
        )
        option = input("Opção: ").strip()
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

    def menu_printers(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "IMPRESSORAS\n1 - Inventário\n2 - Fila\n"
            "3 - Reiniciar Spooler\n0 - Voltar"
        )
        option = input("Opção: ").strip()
        if option == "1":
            self.execute(
                "Inventariando impressoras",
                lambda: self.printers.list_printers(self.host),
                timeout=180,
            )
        elif option == "2":
            self.execute(
                "Consultando filas",
                lambda: self.printers.queue(self.host),
                timeout=180,
            )
        elif option == "3":
            if not self.confirm("Reiniciar o Spooler?"):
                return
            self.execute(
                ActionSpec(
                    "printers.restart_spooler",
                    "Reiniciando Spooler",
                    OperationClass.LIGHT_WRITE,
                    180,
                ),
                lambda: self.printers.restart_spooler(self.host),
            )
        else:
            return
        self.pause()

    def menu_domain(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "DOMÍNIO / GPO\n1 - Status domínio\n"
            "2 - GPResult\n3 - GPUpdate\n0 - Voltar"
        )
        option = input("Opção: ").strip()
        if option == "1":
            self.execute(
                "Verificando domínio",
                lambda: self.domain.status(self.host),
                timeout=180,
            )
        elif option == "2":
            self.execute(
                "Gerando GPResult",
                lambda: self.domain.gpresult(self.host),
                timeout=300,
            )
        elif option == "3":
            self.execute(
                ActionSpec(
                    "domain.gpupdate",
                    "Executando GPUpdate",
                    OperationClass.LIGHT_WRITE,
                    300,
                ),
                lambda: self.domain.gpupdate(self.host),
            )
        else:
            return
        self.pause()

    def menu_storage(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "DISCO / ARMAZENAMENTO / BATERIA\n"
            "1 - Espaço e volumes\n2 - Discos físicos / saúde\n"
            "3 - Bateria\n4 - Perfis por tamanho\n"
            "5 - Estimar limpeza\n0 - Voltar"
        )
        option = input("Opção: ").strip()
        actions = {
            "1": ("Analisando volumes", self.disk.space, OperationClass.READ_ONLY),
            "2": ("Analisando discos físicos", self.storage.physical_disks, OperationClass.READ_ONLY),
            "3": ("Analisando bateria", self.storage.battery, OperationClass.READ_ONLY),
            "4": ("Medindo perfis", self.disk.profile_sizes, OperationClass.HEAVY_READ),
            "5": ("Estimando limpeza", self.disk.cleanup_estimate, OperationClass.HEAVY_READ),
        }
        item = actions.get(option)
        if not item:
            return
        title, function, operation_class = item
        self.execute(
            ActionSpec(
                self._legacy_key(title),
                title,
                operation_class,
                600,
            ),
            lambda: function(self.host),
        )
        self.pause()

    def menu_tools(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "FERRAMENTAS AVANÇADAS\n1 - Certificados\n"
            "2 - Unidades mapeadas\n3 - Compartilhamentos locais\n"
            "4 - Proxy\n5 - Ativação Windows\n"
            "6 - Logons recentes\n0 - Voltar"
        )
        option = input("Opção: ").strip()
        actions = {
            "1": ("Certificados", self.tools.certificates),
            "2": ("Unidades mapeadas", self.tools.mapped_drives),
            "3": ("Compartilhamentos", self.tools.local_shares),
            "4": ("Proxy", self.tools.proxy),
            "5": ("Ativação", self.tools.activation),
            "6": ("Logons", self.tools.logons),
        }
        item = actions.get(option)
        if not item:
            return
        title, function = item
        self.execute(title, lambda: function(self.host), timeout=240)
        self.pause()

    def menu_sysinternals(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "SYSINTERNALS (opcional)\n1 - Ver disponibilidade\n"
            "2 - Autoruns\n3 - ProcDump de processo\n"
            "4 - Procurar handle/arquivo bloqueado\n"
            "5 - Sigcheck de arquivo\n0 - Voltar"
        )
        option = input("Opção: ").strip()
        if option == "1":
            self.execute(
                "Inventariando Sysinternals",
                lambda: self.sysinternals.inventory(self.host),
                timeout=120,
            )
        elif option == "2":
            self.execute(
                "Executando Autorunsc",
                lambda: self.sysinternals.autoruns(self.host),
                timeout=600,
            )
        elif option == "3":
            try:
                process = validate_process_name(
                    input("Processo: ").strip()
                )
            except ValueError as exc:
                print(exc)
                self.pause()
                return
            self.execute(
                ActionSpec(
                    "sysinternals.procdump",
                    "Capturando dump",
                    OperationClass.HEAVY_WRITE,
                    900,
                    True,
                ),
                lambda: self.sysinternals.capture_dump(
                    self.host,
                    process,
                ),
            )
        elif option == "4":
            value = input("Nome/caminho do arquivo: ").strip()
            if value:
                self.execute(
                    "Procurando handles",
                    lambda: self.sysinternals.handle_search(
                        self.host,
                        value,
                    ),
                    timeout=300,
                )
        elif option == "5":
            try:
                value = validate_windows_path(
                    input("Caminho do arquivo: ").strip()
                )
            except ValueError as exc:
                print(exc)
                self.pause()
                return
            self.execute(
                "Validando assinatura/hash",
                lambda: self.sysinternals.sigcheck(
                    self.host,
                    value,
                ),
                timeout=300,
            )
        else:
            return
        self.pause()

    def collect_diagnostic(self) -> None:
        if not self.require_host():
            return
        self.clear()
        result = self.execute(
            ActionSpec(
                "diagnostic.package",
                "Gerando pacote de diagnóstico",
                OperationClass.HEAVY_READ,
                900,
            ),
            lambda: self.packager.collect(self.host),
        )
        if result and result.success:
            print(
                "\nO pacote foi salvo em reports/diagnostics."
            )
        self.pause()

    def menu_system(self) -> None:
        if not self.require_host():
            return
        self.clear()
        print(
            "ENERGIA / PROCESSOS / SERVIÇOS\n"
            "1 - Processos\n2 - Serviços\n"
            "3 - Reiniciar estação\n4 - Desligar estação\n"
            "5 - Enviar mensagem\n0 - Voltar"
        )
        option = input("Opção: ").strip()
        if option == "1":
            self.execute(
                "Listando processos",
                lambda: self.system.list_processes(self.host),
                timeout=180,
            )
        elif option == "2":
            self.execute(
                "Listando serviços",
                lambda: self.system.list_services(self.host),
                timeout=180,
            )
        elif option == "3":
            if not self.confirm(f"REINICIAR {self.host}?"):
                return
            self.execute(
                ActionSpec(
                    "system.restart",
                    "Agendando reinicialização",
                    OperationClass.DISRUPTIVE,
                    120,
                    True,
                ),
                lambda: self.system.restart(self.host),
            )
        elif option == "4":
            if not self.confirm(f"DESLIGAR {self.host}?"):
                return
            self.execute(
                ActionSpec(
                    "system.shutdown",
                    "Agendando desligamento",
                    OperationClass.DISRUPTIVE,
                    120,
                    True,
                ),
                lambda: self.system.shutdown(self.host),
            )
        elif option == "5":
            message = input("Mensagem: ").strip()
            if not message:
                return
            self.execute(
                ActionSpec(
                    "system.message",
                    "Enviando mensagem",
                    OperationClass.LIGHT_WRITE,
                    120,
                ),
                lambda: self.system.send_message(
                    self.host,
                    message,
                ),
            )
        else:
            return
        self.pause()

    def run(self) -> None:
        raise NotImplementedError(
            "ConsoleBase é uma base da interface v5."
        )
