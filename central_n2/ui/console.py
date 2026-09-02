from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from core.executor import RemoteExecutor
from core.result import CommandResult
from modules.batch import BatchRunner
from modules.diagnostics import DiagnosticsModule
from modules.glpi import GLPIModule
from modules.network import NetworkModule
from modules.software import SoftwareModule
from modules.system import SystemModule


class ConsoleUI:
    def __init__(self, executor: RemoteExecutor, settings_path: Path) -> None:
        self.executor = executor
        self.settings_path = settings_path
        self.diag = DiagnosticsModule(executor)
        self.network = NetworkModule(executor)
        self.system = SystemModule(executor)
        self.software = SoftwareModule(executor, settings_path)
        self.glpi = GLPIModule(executor, settings_path)
        self.batch = BatchRunner(max_workers=5)
        self.host: str | None = None
        self.preflight: dict[str, Any] | None = None

    @staticmethod
    def clear() -> None:
        os.system("cls" if os.name == "nt" else "clear")

    @staticmethod
    def pause() -> None:
        input("\nPressione ENTER para continuar...")

    @staticmethod
    def confirm(text: str) -> bool:
        return input(f"\n⚠ {text} [digite SIM]: ").strip().upper() == "SIM"

    def select_host(self) -> None:
        self.clear()
        host = input("Hostname ou IP do computador alvo: ").strip()
        if not host:
            return
        print("\nExecutando pré-flight...")
        self.host = host
        self.preflight = self.diag.preflight(host)
        self._show_preflight()
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
            print(f"Uptime:     {system.get('Uptime') or '-'}")

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

    def run(self) -> None:
        while True:
            self.clear()
            print("╔════════════════════════════════════════════════════╗")
            print("║           CENTRAL REMOTA DE MANUTENÇÃO N2         ║")
            print("╚════════════════════════════════════════════════════╝")
            print(f"\nAlvo atual: {self.host or 'nenhum'}")
            print("\n[1] Selecionar / alterar computador")
            print("[2] Diagnóstico")
            print("[3] Usuários, processos e serviços")
            print("[4] Rede")
            print("[5] Programas")
            print("[6] GLPI Agent")
            print("[7] Políticas / mensagens")
            print("[8] Energia")
            print("[9] Ações em lote")
            print("[0] Sair")
            op = input("\nOpção: ").strip()
            handlers = {
                "1": self.select_host,
                "2": self.menu_diagnostics,
                "3": self.menu_system,
                "4": self.menu_network,
                "5": self.menu_software,
                "6": self.menu_glpi,
                "7": self.menu_policy,
                "8": self.menu_power,
                "9": self.menu_batch,
            }
            if op == "0":
                return
            handler = handlers.get(op)
            if handler:
                handler()

    def menu_diagnostics(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"DIAGNÓSTICO — {self.host}")
            print("1 - Pré-flight completo\n2 - Informações do sistema\n3 - Eventos críticos (24h)\n4 - Reinicialização pendente\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1":
                self.preflight = self.diag.preflight(self.host); self._show_preflight(); self.pause()
            elif op == "2": self.show_result(self.diag.system_info(self.host)); self.pause()
            elif op == "3": self.show_result(self.diag.event_errors(self.host)); self.pause()
            elif op == "4": self.show_result(self.diag.pending_reboot(self.host)); self.pause()

    def menu_system(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"SISTEMA — {self.host}")
            print("1 - Sessões de usuário\n2 - Processos\n3 - Finalizar processo\n4 - Serviços\n5 - Iniciar serviço\n6 - Parar serviço\n7 - Reiniciar serviço\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.system.sessions(self.host)
            elif op == "2": result = self.system.list_processes(self.host)
            elif op == "3":
                name = input("Nome do processo (sem .exe): ").strip()
                if not name or not self.confirm(f"Finalizar o processo '{name}' em {self.host}?"): continue
                result = self.system.kill_process(self.host, name)
            elif op == "4": result = self.system.list_services(self.host)
            elif op in {"5","6","7"}:
                name = input("Nome do serviço: ").strip()
                action = {"5":"start","6":"stop","7":"restart"}[op]
                if action in {"stop","restart"} and not self.confirm(f"{action.upper()} no serviço '{name}'?"): continue
                result = self.system.service_action(self.host, name, action)
            else: continue
            self.show_result(result); self.pause()

    def menu_network(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"REDE — {self.host}")
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
                dest = input("Destino: ").strip(); port = int(input("Porta: ").strip())
                result = self.network.test_tcp(self.host, dest, port)
            else: continue
            self.show_result(result); self.pause()

    def menu_software(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"PROGRAMAS — {self.host}")
            print("1 - Listar instalados\n2 - Verificar Winget\n3 - Instalar do catálogo\n4 - Atualizar do catálogo\n5 - Remover do catálogo\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.software.list_installed(self.host)
            elif op == "2": result = self.software.winget_available(self.host)
            elif op in {"3","4","5"}:
                print("\nCatálogo:")
                for key, item in self.software.catalog.items(): print(f"  {key}: {item['name']}")
                key = input("Chave: ").strip()
                if op == "3": result = self.software.install_catalog_item(self.host, key)
                elif op == "4": result = self.software.upgrade_catalog_item(self.host, key)
                else:
                    if not self.confirm(f"Remover '{key}' de {self.host}?"): continue
                    result = self.software.uninstall_catalog_item(self.host, key)
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

    def menu_policy(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"POLÍTICAS / MENSAGENS — {self.host}")
            print("1 - GPUpdate /force\n2 - Enviar mensagem\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1": result = self.system.gpupdate(self.host)
            elif op == "2": result = self.system.send_message(self.host, input("Mensagem: "))
            else: continue
            self.show_result(result); self.pause()

    def menu_power(self) -> None:
        if not self._require_host(): return
        while True:
            self.clear(); print(f"ENERGIA — {self.host}")
            print("1 - Reiniciar\n2 - Desligar\n3 - Cancelar desligamento/reinicialização agendada\n0 - Voltar")
            op = input("Opção: ").strip()
            if op == "0": return
            if op == "1":
                delay = int(input("Atraso em segundos [0]: ").strip() or "0")
                if not self.confirm(f"REINICIAR {self.host}?"): continue
                result = self.system.restart(self.host, delay)
            elif op == "2":
                delay = int(input("Atraso em segundos [0]: ").strip() or "0")
                if not self.confirm(f"DESLIGAR {self.host}?"): continue
                result = self.system.shutdown(self.host, delay)
            elif op == "3": result = self.system.abort_shutdown(self.host)
            else: continue
            self.show_result(result); self.pause()

    def menu_batch(self) -> None:
        self.clear(); print("AÇÕES EM LOTE")
        print("Informe hostnames separados por vírgula ou um arquivo .txt com um host por linha.")
        raw = input("Hosts/arquivo: ").strip()
        path = Path(raw)
        if path.is_file(): hosts = path.read_text(encoding="utf-8").splitlines()
        else: hosts = [h.strip() for h in raw.split(",")]
        print("\n1 - Pré-flight/ping\n2 - GPUpdate /force\n3 - Flush DNS\n4 - Forçar inventário GLPI\n0 - Voltar")
        op = input("Opção: ").strip()
        if op == "0": return
        actions = {
            "1": self.executor.ping,
            "2": self.system.gpupdate,
            "3": self.network.flush_dns,
            "4": self.glpi.force_inventory,
        }
        action = actions.get(op)
        if not action: return
        results = self.batch.run(hosts, action)
        print("\nRESULTADO")
        for item in results:
            print(f"{item['host']:<25} {'OK' if item['success'] else 'FALHA':<6} {item.get('stderr') or item.get('stdout','')[:80]}")
        self.pause()
