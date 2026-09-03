from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from core.executor import RemoteExecutor
from core.result import CommandResult
from modules.batch import BatchRunner
from modules.compliance import evaluate_compliance
from modules.diagnostic_package import DiagnosticPackageModule
from modules.diagnostics import DiagnosticsModule
from modules.disk import DiskModule
from modules.domain import DomainModule
from modules.glpi import GLPIModule
from modules.health import HealthModule, calculate_health_score
from modules.network import NetworkModule
from modules.printers import PrintersModule
from modules.security import SecurityModule
from modules.software import SoftwareModule
from modules.system import SystemModule
from modules.updates import UpdatesModule
from modules.users_profiles import UsersProfilesModule


class ConsoleUI:
    def __init__(self, executor: RemoteExecutor, settings_path: Path) -> None:
        self.executor = executor
        self.settings_path = settings_path
        self.settings = json.loads(settings_path.read_text(encoding="utf-8"))
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
        self.batch = BatchRunner(max_workers=5)
        self.packager = DiagnosticPackageModule(executor, settings_path.parent.parent / "reports" / "diagnostics")
        self.host: str | None = None
        self.preflight: dict[str, Any] | None = None
        self.health_snapshot: dict[str, Any] | None = None

    @staticmethod
    def clear() -> None:
        os.system("cls" if os.name == "nt" else "clear")

    @staticmethod
    def pause() -> None:
        input("\nPressione ENTER para continuar...")

    @staticmethod
    def confirm(text: str) -> bool:
        return input(f"\n⚠ {text} [digite SIM]: ").strip().upper() == "SIM"

    def _require_host(self) -> bool:
        if self.host:
            return True
        print("Selecione um computador alvo primeiro.")
        self.pause()
        return False

    @staticmethod
    def show_result(result: CommandResult) -> None:
        print(f"\n{'✓ SUCESSO' if result.success else '✗ FALHA'} [{result.transport}]")
        if result.data is not None:
            print(json.dumps(result.data, indent=2, ensure_ascii=False, default=str))
        elif result.stdout:
            print(result.stdout)
        if result.stderr:
            print(f"\nErro: {result.stderr}")
        print(f"\nDuração: {result.duration_ms} ms")

    def select_host(self) -> None:
        self.clear()
        host = input("Hostname ou IP do computador alvo: ").strip()
        if not host:
            return
        print("\nExecutando pré-flight e análise de saúde...")
        self.host = host
        self.preflight = self.diag.preflight(host)
        health_result = self.health.snapshot(host)
        self.health_snapshot = health_result.data if health_result.success and isinstance(health_result.data, dict) else None
        self._show_preflight()
        if self.health_snapshot:
            self._show_health(self.health_snapshot)
        self.pause()

    def _show_preflight(self) -> None:
        p = self.preflight or {}
        system = p.get("system") or {}
        print("\n=== PRÉ-FLIGHT ===")
        print(f"Host:       {p.get('host', '-')}")
        print(f"IP:         {p.get('ip') or 'não resolvido'}")
        print(f"Online:     {'OK' if p.get('online') else 'FALHA'}")
        print(f"Admin$:     {'OK' if p.get('admin_share') else 'FALHA'}")
        print(f"WinRM:      {'OK' if p.get('winrm') else 'indisponível / fallback PsExec'}")
        if system:
            print(f"Usuário:    {system.get('User') or '-'}")
            print(f"SO:         {system.get('OS') or '-'}")
            print(f"Fabricante: {system.get('Manufacturer') or '-'}")
            print(f"Modelo:     {system.get('Model') or '-'}")
            print(f"Serial:     {system.get('Serial') or '-'}")

    def _show_health(self, snapshot: dict[str, Any]) -> None:
        baseline = self.settings.get("compliance", {})
        health = calculate_health_score(
            snapshot,
            min_free_disk_percent=int(baseline.get("min_disk_free_percent", 15)),
            max_uptime_days=int(baseline.get("max_uptime_days", 30)),
        )
        print("\n=== SAÚDE DA ESTAÇÃO ===")
        print(f"Score:      {health['score']}/100")
        print(f"CPU:        {snapshot.get('CPUPercent', '-')}%")
        print(f"RAM:        {snapshot.get('RAMUsedPercent', '-')}% usada")
        print(f"Disco C:    {snapshot.get('DiskFreeGB', '-')} GB livres ({snapshot.get('DiskFreePercent', '-')}%)")
        print(f"Uptime:     {snapshot.get('UptimeDays', '-')} dias")
        if health["findings"]:
            print("\nProblemas encontrados:")
            for item in health["findings"]:
                print(f" - [{item['severity'].upper()}] {item['message']}")
        else:
            print("\nNenhum desvio relevante encontrado pelo baseline atual.")

    def run(self) -> None:
        while True:
            self.clear()
            print("╔════════════════════════════════════════════════════╗")
            print("║             CENTRAL DE MANUTENÇÃO N2              ║")
            print("║              ESTAÇÕES DE TRABALHO                 ║")
            print("╚════════════════════════════════════════════════════╝")
            print(f"\nAlvo atual: {self.host or 'nenhum'}")
            print("\n[1] Selecionar / alterar computador")
            print("[2] Visão geral / Saúde")
            print("[3] Diagnóstico / Hardware")
            print("[4] Usuários e perfis")
            print("[5] Rede")
            print("[6] Software")
            print("[7] Processos e serviços")
            print("[8] Windows Update")
            print("[9] Segurança")
            print("[10] Domínio / GPO")
            print("[11] Impressoras")
            print("[12] GLPI Agent")
            print("[13] Disco / limpeza")
            print("[14] Eventos")
            print("[15] Energia / mensagens")
            print("[16] Coletar diagnóstico")
            print("[17] Compliance")
            print("[18] Ações em lote")
            print("[0] Sair")
            op = input("\nOpção: ").strip()
            handlers = {
                "1": self.select_host,
                "2": self.menu_health,
                "3": self.menu_diagnostics,
                "4": self.menu_users,
                "5": self.menu_network,
                "6": self.menu_software,
                "7": self.menu_system,
                "8": self.menu_updates,
                "9": self.menu_security,
                "10": self.menu_domain,
                "11": self.menu_printers,
                "12": self.menu_glpi,
                "13": self.menu_disk,
                "14": self.menu_events,
                "15": self.menu_power,
                "16": self.collect_diagnostic,
                "17": self.menu_compliance,
                "18": self.menu_batch,
            }
            if op == "0":
                return
            handler = handlers.get(op)
            if handler:
                handler()

    def menu_health(self) -> None:
        if not self._require_host(): return
        self.clear(); print(f"SAÚDE — {self.host}\n")
        result = self.health.snapshot(self.host)
        if result.success and isinstance(result.data, dict):
            self.health_snapshot = result.data
            self._show_health(result.data)
        else:
            self.show_result(result)
        self.pause()

    def menu_diagnostics(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"DIAGNÓSTICO — {self.host}")
            print("1 - Pré-flight completo\n2 - Informações do sistema/hardware\n3 - Reinicialização pendente\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": self.preflight = self.diag.preflight(self.host); self._show_preflight(); self.pause()
            elif op == "2": self.show_result(self.diag.system_info(self.host)); self.pause()
            elif op == "3": self.show_result(self.diag.pending_reboot(self.host)); self.pause()

    def menu_users(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"USUÁRIOS E PERFIS — {self.host}")
            print("1 - Sessões\n2 - Administradores locais\n3 - Perfis e tamanho\n4 - Remover perfil por SID\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.system.sessions(self.host)
            elif op == "2": result = self.users.local_admins(self.host)
            elif op == "3": result = self.users.profiles(self.host)
            elif op == "4":
                sid = input("SID do perfil: ").strip()
                if not sid or not self.confirm(f"REMOVER definitivamente o perfil {sid} de {self.host}?"): continue
                result = self.users.remove_profile(self.host, sid)
            else: continue
            self.show_result(result); self.pause()

    def menu_network(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"REDE DA ESTAÇÃO — {self.host}")
            print("1 - Interfaces\n2 - IP / Gateway / DNS\n3 - Renovar DHCP\n4 - Flush DNS\n5 - Reset Winsock\n6 - Reset TCP/IP\n7 - Desabilitar Wi-Fi\n8 - Habilitar Wi-Fi\n9 - Tabela ARP\n10 - Conexões TCP\n11 - Testar porta TCP\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.network.adapters(self.host)
            elif op == "2": result = self.network.ip_configuration(self.host)
            elif op == "3": result = self.network.renew_dhcp(self.host)
            elif op == "4": result = self.network.flush_dns(self.host)
            elif op == "5":
                if not self.confirm("Resetar Winsock? Pode exigir reinicialização."): continue
                result = self.network.reset_winsock(self.host)
            elif op == "6":
                if not self.confirm("Resetar a pilha TCP/IP? Pode interromper a conectividade."): continue
                result = self.network.reset_tcpip(self.host)
            elif op == "7":
                if not self.confirm("Desabilitar Wi-Fi remotamente?"): continue
                result = self.network.wifi(self.host, False)
            elif op == "8": result = self.network.wifi(self.host, True)
            elif op == "9": result = self.network.arp_table(self.host)
            elif op == "10": result = self.network.connections(self.host)
            elif op == "11":
                dest = input("Destino: ").strip()
                try: port = int(input("Porta: ").strip())
                except ValueError: continue
                result = self.network.test_tcp(self.host, dest, port)
            else: continue
            self.show_result(result); self.pause()

    def menu_software(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"SOFTWARE — {self.host}")
            print("1 - Listar instalados\n2 - Verificar Winget\n3 - Instalar do catálogo\n4 - Atualizar do catálogo\n5 - Remover do catálogo\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.software.list_installed(self.host)
            elif op == "2": result = self.software.winget_available(self.host)
            elif op in {"3","4","5"}:
                for key, item in self.software.catalog.items(): print(f"  {key}: {item['name']}")
                key = input("Chave: ").strip()
                if op == "3": result = self.software.install_catalog_item(self.host, key)
                elif op == "4": result = self.software.upgrade_catalog_item(self.host, key)
                else:
                    if not self.confirm(f"Remover '{key}' de {self.host}?"): continue
                    result = self.software.uninstall_catalog_item(self.host, key)
            else: continue
            self.show_result(result); self.pause()

    def menu_system(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"PROCESSOS E SERVIÇOS — {self.host}")
            print("1 - Processos por CPU\n2 - Finalizar processo\n3 - Serviços\n4 - Iniciar serviço\n5 - Parar serviço\n6 - Reiniciar serviço\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.system.list_processes(self.host)
            elif op == "2":
                name = input("Processo (sem .exe): ").strip()
                if not name or not self.confirm(f"Finalizar '{name}'?"): continue
                result = self.system.kill_process(self.host, name)
            elif op == "3": result = self.system.list_services(self.host)
            elif op in {"4","5","6"}:
                name = input("Nome do serviço: ").strip()
                action = {"4":"start","5":"stop","6":"restart"}[op]
                if action != "start" and not self.confirm(f"{action.upper()} no serviço '{name}'?"): continue
                result = self.system.service_action(self.host, name, action)
            else: continue
            self.show_result(result); self.pause()

    def menu_updates(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"WINDOWS UPDATE — {self.host}")
            print("1 - Status / histórico / pendentes\n2 - Iniciar busca\n3 - Resetar componentes do Windows Update\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.updates.status(self.host)
            elif op == "2": result = self.updates.trigger_scan(self.host)
            elif op == "3":
                if not self.confirm("Resetar componentes do Windows Update? Os serviços serão reiniciados."): continue
                result = self.updates.reset_components(self.host)
            else: continue
            self.show_result(result); self.pause()

    def menu_security(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"SEGURANÇA — {self.host}")
            print("1 - Defender / Firewall / BitLocker / TPM / Secure Boot\n2 - Ameaças recentes\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            result = self.security.status(self.host) if op == "1" else self.security.threats(self.host) if op == "2" else None
            if result: self.show_result(result); self.pause()

    def menu_domain(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"DOMÍNIO / GPO — {self.host}")
            print("1 - Status do domínio / secure channel / hora\n2 - GPResult computador\n3 - GPUpdate /force\n4 - Reparar secure channel\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.domain.status(self.host)
            elif op == "2": result = self.domain.gpresult(self.host)
            elif op == "3": result = self.system.gpupdate(self.host)
            elif op == "4":
                if not self.confirm("Tentar reparar o secure channel com o domínio?"): continue
                result = self.domain.repair_secure_channel(self.host)
            else: continue
            self.show_result(result); self.pause()

    def menu_printers(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"IMPRESSORAS — {self.host}")
            print("1 - Listar impressoras\n2 - Ver filas\n3 - Reiniciar Spooler\n4 - Limpar todas as filas\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.printers.list(self.host)
            elif op == "2": result = self.printers.queue(self.host)
            elif op == "3":
                if not self.confirm("Reiniciar o serviço Spooler?"): continue
                result = self.printers.restart_spooler(self.host)
            elif op == "4":
                if not self.confirm("Remover TODOS os trabalhos das filas de impressão?"): continue
                result = self.printers.clear_queue(self.host)
            else: continue
            self.show_result(result); self.pause()

    def menu_glpi(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"GLPI AGENT — {self.host}")
            print("1 - Status\n2 - Instalar / reparar\n3 - Reiniciar serviço\n4 - Forçar inventário\n5 - Log recente\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.glpi.status(self.host)
            elif op == "2": result = self.glpi.install_or_repair(self.host)
            elif op == "3": result = self.glpi.restart_service(self.host)
            elif op == "4": result = self.glpi.force_inventory(self.host)
            elif op == "5": result = self.glpi.recent_log(self.host)
            else: continue
            self.show_result(result); self.pause()

    def menu_disk(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"DISCO / LIMPEZA — {self.host}")
            print("1 - Uso do disco C:\n2 - Tamanho dos perfis\n3 - Estimar espaço recuperável\n4 - Limpeza segura (TEMP + Lixeira)\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.disk.usage(self.host)
            elif op == "2": result = self.disk.top_user_profiles(self.host)
            elif op == "3": result = self.disk.cleanup_estimate(self.host)
            elif op == "4":
                if not self.confirm("Executar limpeza segura de TEMP e Lixeira? Downloads não serão apagados."): continue
                result = self.disk.cleanup_safe(self.host)
            else: continue
            self.show_result(result); self.pause()

    def menu_events(self) -> None:
        if not self._require_host(): return
        self.clear(); print(f"EVENTOS CRÍTICOS — {self.host}")
        self.show_result(self.diag.event_errors(self.host)); self.pause()

    def menu_power(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"ENERGIA / MENSAGENS — {self.host}")
            print("1 - Enviar mensagem\n2 - Reiniciar\n3 - Desligar\n4 - Cancelar shutdown agendado\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.system.send_message(self.host, input("Mensagem: "))
            elif op in {"2","3"}:
                try: delay = int(input("Atraso em segundos [0]: ").strip() or "0")
                except ValueError: continue
                action = "REINICIAR" if op == "2" else "DESLIGAR"
                if not self.confirm(f"{action} {self.host}?"): continue
                result = self.system.restart(self.host, delay) if op == "2" else self.system.shutdown(self.host, delay)
            elif op == "4": result = self.system.abort_shutdown(self.host)
            else: continue
            self.show_result(result); self.pause()

    def collect_diagnostic(self) -> None:
        if not self._require_host(): return
        self.clear(); print(f"COLETA DE DIAGNÓSTICO — {self.host}\n")
        try:
            path = self.packager.collect(self.host)
            print(f"✓ Relatório salvo em: {path}")
        except Exception as exc:
            print(f"✗ Falha: {exc}")
        self.pause()

    def menu_compliance(self) -> None:
        if not self._require_host(): return
        self.clear(); print(f"COMPLIANCE — {self.host}\n")
        result = self.health.snapshot(self.host)
        if not result.success or not isinstance(result.data, dict):
            self.show_result(result); self.pause(); return
        report = evaluate_compliance(result.data, self.settings.get("compliance", {}))
        print(f"Score: {report['score']}% ({report['compliant']}/{report['total']} controles)")
        for item in report["items"]:
            mark = "✓" if item["compliant"] else "✗"
            print(f"{mark} {item['label']}: atual={item['actual']} esperado={item['expected']}")
        self.pause()

    def menu_batch(self) -> None:
        self.clear(); print("AÇÕES EM LOTE")
        raw = input("Hosts separados por vírgula ou arquivo .txt: ").strip()
        path = Path(raw)
        hosts = path.read_text(encoding="utf-8").splitlines() if path.is_file() else [h.strip() for h in raw.split(",") if h.strip()]
        print("\n1 - Ping\n2 - GPUpdate /force\n3 - Flush DNS\n4 - Forçar inventário GLPI\n5 - Saúde básica\n0 - Voltar")
        op = input("Opção: ").strip()
        if op == "0": return
        actions = {
            "1": self.executor.ping,
            "2": self.system.gpupdate,
            "3": self.network.flush_dns,
            "4": self.glpi.force_inventory,
            "5": self.health.snapshot,
        }
        action = actions.get(op)
        if not action: return
        results = self.batch.run(hosts, action)
        print("\nRESULTADO")
        for item in results:
            print(f"{item['host']:<25} {'OK' if item['success'] else 'FALHA':<6} {item.get('stderr') or item.get('stdout','')[:80]}")
        self.pause()
