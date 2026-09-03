from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Callable

from core.jobs import ResponsiveJobRunner
from core.result import CommandResult
from modules.compliance import evaluate_compliance
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


class ConsoleUIV3:
    def __init__(self, executor, settings_path: Path) -> None:
        self.executor = executor
        self.settings_path = settings_path
        self.settings = json.loads(settings_path.read_text(encoding="utf-8"))
        ui = self.settings.get("ui", {})
        self.jobs = ResponsiveJobRunner(heartbeat_seconds=float(ui.get("heartbeat_seconds", 0.2)))
        self.long_timeout = int(ui.get("long_operation_timeout_seconds", 3600))
        self.host: str | None = None
        self.health_snapshot: dict[str, Any] | None = None

        self.diag = DiagnosticsModule(executor)
        self.health = HealthModule(executor)
        self.network = NetworkModule(executor)
        self.system = SystemModule(executor)
        self.software = SoftwareModule(executor, settings_path)
        self.glpi = GLPIModule(executor, settings_path)
        self.security = SecurityModule(executor)
        self.updates = UpdatesModule(executor)
        self.users = UsersProfilesModule(executor)
        self.printers = PrintersModule(executor)
        self.domain = DomainModule(executor)
        self.disk = DiskModule(executor)
        self.packager = DiagnosticPackageModule(executor, settings_path.parent.parent / "reports" / "diagnostics")
        self.repair = RepairModule(executor)
        self.devices = DevicesModule(executor)
        self.performance = PerformanceModule(executor)
        self.startup = StartupModule(executor)
        self.crashes = CrashesModule(executor)
        self.tasks = TasksModule(executor)
        self.storage = StorageModule(executor)
        self.sysinternals = SysinternalsModule(executor, self.settings.get("sysinternals_dir", r"C:\Sysinternals"))
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
    def show_result(result: CommandResult) -> None:
        status = "✓ SUCESSO" if result.success else "✗ FALHA"
        print(f"\n{status} [{result.transport}] — {result.duration_ms} ms")
        if result.data is not None:
            print(json.dumps(result.data, indent=2, ensure_ascii=False, default=str))
        elif result.stdout:
            print(result.stdout)
        if result.stderr:
            print(f"\nErro: {result.stderr}")

    def execute(self, label: str, func: Callable[[], CommandResult], *, timeout: int | None = None) -> CommandResult | None:
        print(f"\n▶ {label}")
        try:
            result = self.jobs.run(label, func, timeout=timeout or self.long_timeout)
        except TimeoutError as exc:
            print(f"\n✗ TIMEOUT: {exc}")
            return None
        except Exception as exc:
            print(f"\n✗ ERRO NÃO TRATADO: {type(exc).__name__}: {exc}")
            return None
        self.show_result(result)
        return result

    def select_host(self) -> None:
        self.clear()
        host = input("Hostname ou IP da estação: ").strip()
        if not host:
            return
        self.host = host
        result = self.execute("Pré-flight da estação", lambda: self.health.snapshot(host), timeout=120)
        if result and result.success and isinstance(result.data, dict):
            self.health_snapshot = result.data
            self.show_health(result.data)
        self.pause()

    def show_health(self, snapshot: dict[str, Any]) -> None:
        base = self.settings.get("compliance", {})
        health = calculate_health_score(snapshot, min_free_disk_percent=int(base.get("min_disk_free_percent", 15)), max_uptime_days=int(base.get("max_uptime_days", 30)))
        print("\n=== SAÚDE DA ESTAÇÃO ===")
        print(f"Score: {health['score']}/100 | CPU: {snapshot.get('CPUPercent','-')}% | RAM: {snapshot.get('RAMUsedPercent','-')}% | Disco livre: {snapshot.get('DiskFreePercent','-')}% | Uptime: {snapshot.get('UptimeDays','-')} dias")
        for item in health["findings"]:
            print(f" - [{item['severity'].upper()}] {item['message']}")
        if not health["findings"]:
            print("Nenhum desvio relevante encontrado.")

    def run(self) -> None:
        try:
            while True:
                self.clear()
                print("╔════════════════════════════════════════════════════╗")
                print("║        CENTRAL N2 WORKSTATION — V3 RESPONSIVA     ║")
                print("╚════════════════════════════════════════════════════╝")
                print(f"\nAlvo: {self.host or 'nenhum'}")
                print("\n[1] Selecionar estação")
                print("[2] Saúde / Compliance")
                print("[3] Performance")
                print("[4] Reparo do Windows")
                print("[5] Hardware / Drivers / Dispositivos")
                print("[6] Inicialização / Tarefas")
                print("[7] Crashes / BSOD")
                print("[8] Segurança")
                print("[9] Rede")
                print("[10] Usuários / Perfis")
                print("[11] Software / GLPI")
                print("[12] Impressoras")
                print("[13] Domínio / GPO")
                print("[14] Disco / Armazenamento / Bateria")
                print("[15] Ferramentas avançadas")
                print("[16] Sysinternals")
                print("[17] Pacote de diagnóstico")
                print("[18] Energia / Processos / Serviços")
                print("[0] Sair")
                op = input("\nOpção: ").strip()
                handlers = {
                    "1": self.select_host, "2": self.menu_health, "3": self.menu_performance,
                    "4": self.menu_repair, "5": self.menu_devices, "6": self.menu_startup_tasks,
                    "7": self.menu_crashes, "8": self.menu_security, "9": self.menu_network,
                    "10": self.menu_users, "11": self.menu_software_glpi, "12": self.menu_printers,
                    "13": self.menu_domain, "14": self.menu_storage, "15": self.menu_tools,
                    "16": self.menu_sysinternals, "17": self.collect_diagnostic, "18": self.menu_system,
                }
                if op == "0": return
                if op in handlers: handlers[op]()
        finally:
            self.jobs.shutdown()

    def menu_health(self) -> None:
        if not self.require_host(): return
        self.clear(); result=self.execute("Coletando saúde", lambda:self.health.snapshot(self.host), timeout=120)
        if result and result.success and isinstance(result.data,dict):
            self.health_snapshot=result.data; self.show_health(result.data)
            comp=evaluate_compliance(result.data,self.settings.get("compliance",{}))
            print(f"\nCompliance: {comp['score']}/100")
            for f in comp['findings']: print(f" - {f}")
        self.pause()

    def menu_performance(self) -> None:
        if not self.require_host(): return
        self.clear(); print("PERFORMANCE\n1 - Amostragem rápida (8s)\n2 - Amostragem detalhada (20s)\n0 - Voltar")
        op=input("Opção: ").strip()
        if op=="1": self.execute("Amostrando CPU/RAM/disco/rede",lambda:self.performance.snapshot(self.host,8,1),timeout=120)
        elif op=="2": self.execute("Amostrando performance detalhada",lambda:self.performance.snapshot(self.host,20,1),timeout=180)
        else: return
        self.pause()

    def menu_repair(self) -> None:
        if not self.require_host(): return
        actions={"1":("SFC /scannow",self.repair.sfc_scan),"2":("DISM CheckHealth",self.repair.dism_checkhealth),"3":("DISM ScanHealth",self.repair.dism_scanhealth),"4":("DISM RestoreHealth",self.repair.dism_restorehealth),"5":("Analisar Component Store",self.repair.component_store),"6":("Limpar Component Store",self.repair.component_cleanup),"7":("CHKDSK /scan",self.repair.chkdsk_scan),"8":("Verificar WMI repository",self.repair.repository_consistency)}
        self.clear(); print("REPARO DO WINDOWS\n1 - SFC\n2 - DISM CheckHealth\n3 - DISM ScanHealth\n4 - DISM RestoreHealth\n5 - Analisar Component Store\n6 - Limpar Component Store\n7 - CHKDSK online\n8 - Verificar WMI\n0 - Voltar")
        op=input("Opção: ").strip(); item=actions.get(op)
        if not item:return
        if op in {"4","6"} and not self.confirm(f"Executar {item[0]} em {self.host}?"): return
        self.execute(item[0],lambda:item[1](self.host),timeout=3600); self.pause()

    def menu_devices(self) -> None:
        if not self.require_host(): return
        self.clear(); print("DISPOSITIVOS / DRIVERS\n1 - Dispositivos com erro\n2 - Drivers\n3 - USB presentes\n4 - Reexaminar dispositivos\n5 - Exportar drivers\n0 - Voltar")
        op=input("Opção: ").strip()
        actions={"1":("Analisando dispositivos com erro",self.devices.problem_devices),"2":("Inventariando drivers",self.devices.drivers),"3":("Inventariando USB",self.devices.usb_devices),"4":("Reexaminando hardware",self.devices.rescan),"5":("Exportando drivers",self.devices.export_drivers)}
        item=actions.get(op)
        if not item:return
        self.execute(item[0],lambda:item[1](self.host),timeout=1800); self.pause()

    def menu_startup_tasks(self) -> None:
        if not self.require_host(): return
        self.clear(); print("INICIALIZAÇÃO / TAREFAS\n1 - Visão de inicialização\n2 - Tarefas agendadas\n3 - Tarefas com falha\n0 - Voltar")
        op=input("Opção: ").strip(); actions={"1":("Analisando inicialização",self.startup.overview),"2":("Listando tarefas",self.tasks.list_tasks),"3":("Localizando tarefas com falha",self.tasks.failed_tasks)}; item=actions.get(op)
        if not item:return
        self.execute(item[0],lambda:item[1](self.host),timeout=240); self.pause()

    def menu_crashes(self) -> None:
        if not self.require_host(): return
        self.clear(); print("CRASHES / BSOD\n1 - BSOD e dumps\n2 - Crashes de aplicações\n0 - Voltar")
        op=input("Opção: ").strip(); actions={"1":("Coletando histórico de BSOD",self.crashes.bsod_history),"2":("Coletando crashes de aplicações",self.crashes.app_crashes)}; item=actions.get(op)
        if not item:return
        self.execute(item[0],lambda:item[1](self.host),timeout=240); self.pause()

    def menu_security(self) -> None:
        if not self.require_host(): return
        self.clear(); print("SEGURANÇA\n1 - Postura de segurança\n2 - Ameaças recentes\n0 - Voltar")
        op=input("Opção: ").strip()
        if op=="1": self.execute("Coletando postura de segurança",lambda:self.security.posture(self.host),timeout=180)
        elif op=="2": self.execute("Consultando ameaças",lambda:self.security.threats(self.host),timeout=180)
        else:return
        self.pause()

    def menu_network(self) -> None:
        if not self.require_host(): return
        self.clear(); print("REDE\n1 - Interfaces\n2 - IP/Gateway/DNS\n3 - ARP\n4 - Conexões TCP\n5 - Flush DNS\n6 - Renovar DHCP\n0 - Voltar")
        op=input("Opção: ").strip(); actions={"1":("Interfaces",self.network.adapters),"2":("Configuração IP",self.network.ip_configuration),"3":("Tabela ARP",self.network.arp_table),"4":("Conexões TCP",self.network.connections),"5":("Flush DNS",self.network.flush_dns),"6":("Renovar DHCP",self.network.renew_dhcp)}; item=actions.get(op)
        if not item:return
        self.execute(item[0],lambda:item[1](self.host),timeout=180); self.pause()

    def menu_users(self) -> None:
        if not self.require_host(): return
        self.clear(); print("USUÁRIOS / PERFIS\n1 - Sessões\n2 - Administradores locais\n3 - Perfis e tamanho\n0 - Voltar")
        op=input("Opção: ").strip(); actions={"1":("Sessões",self.system.sessions),"2":("Administradores locais",self.users.local_admins),"3":("Perfis",self.users.profiles)}; item=actions.get(op)
        if not item:return
        self.execute(item[0],lambda:item[1](self.host),timeout=300); self.pause()

    def menu_software_glpi(self) -> None:
        if not self.require_host(): return
        self.clear(); print("SOFTWARE / GLPI\n1 - Software instalado\n2 - Winget\n3 - GLPI status\n4 - Forçar inventário GLPI\n0 - Voltar")
        op=input("Opção: ").strip()
        if op=="1": fn=self.software.list_installed; label="Inventariando software"
        elif op=="2": fn=self.software.winget_available; label="Verificando Winget"
        elif op=="3": fn=self.glpi.status; label="Verificando GLPI"
        elif op=="4": fn=self.glpi.force_inventory; label="Forçando inventário GLPI"
        else:return
        self.execute(label,lambda:fn(self.host),timeout=600); self.pause()

    def menu_printers(self) -> None:
        if not self.require_host(): return
        self.clear(); print("IMPRESSORAS\n1 - Inventário\n2 - Fila\n3 - Reiniciar Spooler\n0 - Voltar")
        op=input("Opção: ").strip()
        if op=="1": self.execute("Inventariando impressoras",lambda:self.printers.list_printers(self.host),timeout=180)
        elif op=="2": self.execute("Consultando filas",lambda:self.printers.queue(self.host),timeout=180)
        elif op=="3" and self.confirm("Reiniciar o Spooler?"): self.execute("Reiniciando Spooler",lambda:self.printers.restart_spooler(self.host),timeout=180)
        else:return
        self.pause()

    def menu_domain(self) -> None:
        if not self.require_host(): return
        self.clear(); print("DOMÍNIO / GPO\n1 - Status domínio\n2 - GPResult\n3 - GPUpdate\n0 - Voltar")
        op=input("Opção: ").strip()
        if op=="1": self.execute("Verificando domínio",lambda:self.domain.status(self.host),timeout=180)
        elif op=="2": self.execute("Gerando GPResult",lambda:self.domain.gpresult(self.host),timeout=300)
        elif op=="3": self.execute("Executando GPUpdate",lambda:self.domain.gpupdate(self.host),timeout=300)
        else:return
        self.pause()

    def menu_storage(self) -> None:
        if not self.require_host(): return
        self.clear(); print("DISCO / ARMAZENAMENTO / BATERIA\n1 - Espaço e volumes\n2 - Discos físicos / saúde\n3 - Bateria\n4 - Perfis por tamanho\n5 - Estimar limpeza\n0 - Voltar")
        op=input("Opção: ").strip()
        actions={"1":("Analisando volumes",self.disk.space),"2":("Analisando discos físicos",self.storage.physical_disks),"3":("Analisando bateria",self.storage.battery),"4":("Medindo perfis",self.disk.profile_sizes),"5":("Estimando limpeza",self.disk.cleanup_estimate)}; item=actions.get(op)
        if not item:return
        self.execute(item[0],lambda:item[1](self.host),timeout=600); self.pause()

    def menu_tools(self) -> None:
        if not self.require_host(): return
        self.clear(); print("FERRAMENTAS AVANÇADAS\n1 - Certificados\n2 - Unidades mapeadas\n3 - Compartilhamentos locais\n4 - Proxy\n5 - Ativação Windows\n6 - Logons recentes\n0 - Voltar")
        op=input("Opção: ").strip(); actions={"1":("Certificados",self.tools.certificates),"2":("Unidades mapeadas",self.tools.mapped_drives),"3":("Compartilhamentos",self.tools.local_shares),"4":("Proxy",self.tools.proxy),"5":("Ativação",self.tools.activation),"6":("Logons",self.tools.logons)}; item=actions.get(op)
        if not item:return
        self.execute(item[0],lambda:item[1](self.host),timeout=240); self.pause()

    def menu_sysinternals(self) -> None:
        if not self.require_host(): return
        self.clear(); print("SYSINTERNALS (opcional)\n1 - Ver disponibilidade\n2 - Autoruns\n3 - ProcDump de processo\n4 - Procurar handle/arquivo bloqueado\n5 - Sigcheck de arquivo\n0 - Voltar")
        op=input("Opção: ").strip()
        if op=="1": self.execute("Inventariando Sysinternals",lambda:self.sysinternals.inventory(self.host),timeout=120)
        elif op=="2": self.execute("Executando Autorunsc",lambda:self.sysinternals.autoruns(self.host),timeout=600)
        elif op=="3":
            p=input("Processo: ").strip();
            if p:self.execute("Capturando dump",lambda:self.sysinternals.capture_dump(self.host,p),timeout=900)
        elif op=="4":
            p=input("Nome/caminho do arquivo: ").strip();
            if p:self.execute("Procurando handles",lambda:self.sysinternals.handle_search(self.host,p),timeout=300)
        elif op=="5":
            p=input("Caminho do arquivo: ").strip();
            if p:self.execute("Validando assinatura/hash",lambda:self.sysinternals.sigcheck(self.host,p),timeout=300)
        else:return
        self.pause()

    def collect_diagnostic(self) -> None:
        if not self.require_host(): return
        self.clear(); result=self.execute("Gerando pacote de diagnóstico",lambda:self.packager.collect(self.host),timeout=900)
        if result: print("\nO pacote foi gerado na pasta reports/diagnostics quando a coleta local foi concluída.")
        self.pause()

    def menu_system(self) -> None:
        if not self.require_host(): return
        self.clear(); print("ENERGIA / PROCESSOS / SERVIÇOS\n1 - Processos\n2 - Serviços\n3 - Reiniciar estação\n4 - Desligar estação\n5 - Enviar mensagem\n0 - Voltar")
        op=input("Opção: ").strip()
        if op=="1": self.execute("Listando processos",lambda:self.system.list_processes(self.host),timeout=180)
        elif op=="2": self.execute("Listando serviços",lambda:self.system.list_services(self.host),timeout=180)
        elif op=="3" and self.confirm(f"REINICIAR {self.host}?"): self.execute("Agendando reinicialização",lambda:self.system.restart(self.host),timeout=120)
        elif op=="4" and self.confirm(f"DESLIGAR {self.host}?"): self.execute("Agendando desligamento",lambda:self.system.shutdown(self.host),timeout=120)
        elif op=="5":
            msg=input("Mensagem: ")
            if msg:self.execute("Enviando mensagem",lambda:self.system.send_message(self.host,msg),timeout=120)
        else:return
        self.pause()
