from __future__ import annotations

import json
from pathlib import Path

from modules.network_nac import NetworkNACModule, PROFILES, SwitchSSHClient


BASE_DIR = Path(__file__).resolve().parent
ALLOWLIST = BASE_DIR / "config" / "oui_allowlist.json"
REPORT_DIR = BASE_DIR / "reports" / "network_nac"


def pause() -> None:
    input("\nPressione ENTER para continuar...")


def load_table_from_file() -> str:
    path = Path(input("Arquivo com saída da tabela MAC: ").strip().strip('"'))
    return path.read_text(encoding="utf-8", errors="replace")


def collect_table_ssh(profile_key: str) -> str:
    profile = PROFILES[profile_key]
    host = input("IP/hostname do switch: ").strip()
    username = input("Usuário SSH: ").strip()
    identity = input("Chave privada SSH (vazio = padrão do usuário): ").strip().strip('"') or None
    client = SwitchSSHClient()
    result = client.run(host, username, [profile.show_mac_command], identity_file=identity)
    if result.returncode != 0:
        raise RuntimeError(result.stderr or result.stdout or "Falha ao consultar switch via SSH")
    return result.stdout


def show_audit(summary: dict) -> None:
    print("\n=== AUDITORIA MAC/OUI ===")
    print(f"Total:          {summary['total']}")
    print(f"Autorizados:    {summary['authorized']}")
    print(f"Não autorizados:{summary['unauthorized']}")
    print("\nMAC                  PORTA           VLAN   FABRICANTE              STATUS / MOTIVO")
    print("-" * 105)
    for item in summary["items"]:
        status = "OK" if item["authorized"] else "BLOQUEAR"
        print(
            f"{item['mac']:<20} {(item['port'] or '-'):<15} {(item['vlan'] or '-')!s:<6} "
            f"{(item['manufacturer'] or '-'):<23} {status:<9} {item['reason']}"
        )


def save_report(summary: dict) -> Path:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    path = REPORT_DIR / "ultimo_relatorio.json"
    path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def select_profile() -> str:
    print("\nPerfis disponíveis:")
    keys = list(PROFILES)
    for index, key in enumerate(keys, start=1):
        print(f"{index} - {PROFILES[key].display_name}")
    selected = int(input("Perfil: ").strip())
    if selected < 1 or selected > len(keys):
        raise ValueError("Perfil inválido")
    return keys[selected - 1]


def generate_plan(module: NetworkNACModule, profile_key: str) -> None:
    profile = PROFILES[profile_key]
    print(f"\nPerfil: {profile.display_name}")
    print(profile.notes)
    if not profile.prefix_acl_supported:
        print("\nEste perfil não gera allowlist OUI automaticamente nesta versão.")
        print("Use a auditoria para identificar MACs e valide ACL exata/modelo antes da implantação.")
        return

    ports_raw = input("Portas DE ACESSO para proteção (separadas por vírgula): ").strip()
    ports = [p.strip() for p in ports_raw.split(",") if p.strip()]
    excluded_raw = input("Portas a EXCLUIR (uplinks/trunks/APs), separadas por vírgula: ").strip()
    excluded = {p.strip().lower() for p in excluded_raw.split(",") if p.strip()}
    ports = [p for p in ports if p.lower() not in excluded]
    acl_id = int(input("ID da ACL [2001]: ").strip() or "2001")

    commands = module.plan_prefix_acl(profile_key, ports, acl_id=acl_id)
    print("\n=== PLANO DE CONFIGURAÇÃO (NÃO APLICADO) ===")
    print("\n".join(commands))

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    plan_path = REPORT_DIR / "ultimo_plano_acl.txt"
    plan_path.write_text("\n".join(commands) + "\n", encoding="utf-8")
    print(f"\nPlano salvo em: {plan_path}")
    print("A aplicação automática foi deliberadamente separada da geração. Valide o plano em bancada primeiro.")


def main() -> None:
    module = NetworkNACModule(ALLOWLIST)
    while True:
        print("\n╔════════════════════════════════════════════════════╗")
        print("║          CENTRAL N2 — NETWORK / NAC INTELBRAS     ║")
        print("╚════════════════════════════════════════════════════╝")
        print("1 - Auditar tabela MAC a partir de arquivo")
        print("2 - Auditar tabela MAC diretamente via SSH")
        print("3 - Gerar plano de ACL por OUI (sem aplicar)")
        print("4 - Mostrar OUIs homologados")
        print("0 - Sair")
        op = input("\nOpção: ").strip()
        if op == "0":
            return
        try:
            if op == "1":
                text = load_table_from_file()
                summary = module.audit_summary(text)
                show_audit(summary)
                print(f"\nRelatório: {save_report(summary)}")
                pause()
            elif op == "2":
                profile = select_profile()
                text = collect_table_ssh(profile)
                summary = module.audit_summary(text)
                show_audit(summary)
                print(f"\nRelatório: {save_report(summary)}")
                pause()
            elif op == "3":
                profile = select_profile()
                generate_plan(module, profile)
                pause()
            elif op == "4":
                for manufacturer, prefixes in module.allowlist.manufacturers.items():
                    print(f"\n{manufacturer}: {', '.join(sorted(prefixes)) or '(nenhum OUI cadastrado)'}")
                pause()
        except Exception as exc:
            print(f"\nERRO: {exc}")
            pause()


if __name__ == "__main__":
    main()
