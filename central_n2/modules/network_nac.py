from __future__ import annotations

import json
import re
import shutil
import subprocess
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable


MAC_RE = re.compile(r"(?i)\b([0-9a-f]{2}(?:[:-][0-9a-f]{2}){5}|[0-9a-f]{4}(?:\.[0-9a-f]{4}){2}|[0-9a-f]{12})\b")
PORT_RE = re.compile(r"(?i)\b((?:ge|gi|gigabitethernet|ethernet|eth|te|xe)[\w./-]+)\b")
VLAN_RE = re.compile(r"\b([1-9]\d{0,3})\b")


def normalize_mac(value: str) -> str:
    raw = re.sub(r"[^0-9A-Fa-f]", "", value).upper()
    if len(raw) != 12:
        raise ValueError(f"MAC inválido: {value!r}")
    return ":".join(raw[i:i + 2] for i in range(0, 12, 2))


def mac_to_intelbras(value: str) -> str:
    raw = normalize_mac(value).replace(":", "")
    return f"{raw[0:4]}.{raw[4:8]}.{raw[8:12]}"


def normalize_prefix(value: str) -> str:
    raw = re.sub(r"[^0-9A-Fa-f]", "", value).upper()
    if len(raw) != 6:
        raise ValueError(f"Prefixo OUI deve possuir 24 bits (6 hex): {value!r}")
    return ":".join(raw[i:i + 2] for i in range(0, 6, 2))


def prefix_to_serie3000_acl(prefix: str) -> tuple[str, str]:
    """Retorna endereço MAC base + wildcard para um OUI /24 na Série 3000."""
    raw = normalize_prefix(prefix).replace(":", "")
    base = f"{raw[0:4]}.{raw[4:6]}00.0000"
    wildcard = "0000.00FF.FFFF"
    return base, wildcard


def is_locally_administered(mac: str) -> bool:
    first_octet = int(normalize_mac(mac).split(":")[0], 16)
    return bool(first_octet & 0x02)


@dataclass(frozen=True)
class MacEntry:
    mac: str
    port: str | None = None
    vlan: int | None = None
    entry_type: str | None = None

    @property
    def oui(self) -> str:
        return ":".join(self.mac.split(":")[:3])


@dataclass(frozen=True)
class Classification:
    mac: str
    oui: str
    port: str | None
    vlan: int | None
    manufacturer: str | None
    authorized: bool
    reason: str


@dataclass(frozen=True)
class SwitchProfile:
    key: str
    display_name: str
    show_mac_command: str
    prefix_acl_supported: bool
    exact_acl_supported: bool
    notes: str


PROFILES: dict[str, SwitchProfile] = {
    "intelbras_serie3000": SwitchProfile(
        key="intelbras_serie3000",
        display_name="Intelbras Série 3000",
        show_mac_command="show mac address-table",
        prefix_acl_supported=True,
        exact_acl_supported=True,
        notes="MAC ACL padrão/estendida com wildcard documentado; valide nomes de interface no equipamento.",
    ),
    "intelbras_s2050g_a": SwitchProfile(
        key="intelbras_s2050g_a",
        display_name="Intelbras S2050G-A / família CLI semelhante",
        show_mac_command="show mac-address-table",
        prefix_acl_supported=False,
        exact_acl_supported=True,
        notes="A documentação confirma MAC ACL e mac-access-group, mas a v1 não gera allowlist OUI por wildcard neste perfil.",
    ),
    "intelbras_s_series": SwitchProfile(
        key="intelbras_s_series",
        display_name="Intelbras S-Series",
        show_mac_command="show mac-address all",
        prefix_acl_supported=False,
        exact_acl_supported=True,
        notes="MAC ACL documentada com host/any; prefixo OUI não é aplicado automaticamente pela v1.",
    ),
}


class OuiAllowlist:
    def __init__(self, path: Path) -> None:
        self.path = path
        payload = json.loads(path.read_text(encoding="utf-8"))
        self.manufacturers: dict[str, set[str]] = {}
        for item in payload.get("manufacturers", []):
            name = str(item["name"]).strip()
            prefixes = {normalize_prefix(p) for p in item.get("prefixes", [])}
            self.manufacturers[name] = prefixes
        self.exact_macs: dict[str, str] = {
            normalize_mac(item["mac"]): str(item.get("name") or "Exceção")
            for item in payload.get("exact_macs", [])
        }

    @property
    def allowed_prefixes(self) -> set[str]:
        result: set[str] = set()
        for prefixes in self.manufacturers.values():
            result.update(prefixes)
        return result

    def manufacturer_for(self, mac: str) -> str | None:
        oui = ":".join(normalize_mac(mac).split(":")[:3])
        for name, prefixes in self.manufacturers.items():
            if oui in prefixes:
                return name
        return None

    def classify(self, entry: MacEntry) -> Classification:
        mac = normalize_mac(entry.mac)
        oui = ":".join(mac.split(":")[:3])
        if mac in self.exact_macs:
            return Classification(mac, oui, entry.port, entry.vlan, self.exact_macs[mac], True, "exceção por MAC exato")
        manufacturer = self.manufacturer_for(mac)
        if manufacturer:
            return Classification(mac, oui, entry.port, entry.vlan, manufacturer, True, "OUI homologado")
        if is_locally_administered(mac):
            return Classification(mac, oui, entry.port, entry.vlan, None, False, "MAC local/randomizado; não corresponde a OUI global homologado")
        return Classification(mac, oui, entry.port, entry.vlan, None, False, "OUI não homologado")


class MacTableParser:
    """Parser tolerante a formatos comuns das famílias Intelbras."""

    @staticmethod
    def parse(text: str) -> list[MacEntry]:
        entries: list[MacEntry] = []
        seen: set[tuple[str, str | None, int | None]] = set()
        for line in text.splitlines():
            match = MAC_RE.search(line)
            if not match:
                continue
            try:
                mac = normalize_mac(match.group(1))
            except ValueError:
                continue

            port_match = PORT_RE.search(line)
            port = port_match.group(1) if port_match else None
            lowered = line.lower()
            entry_type = None
            if "dynamic" in lowered or "dinamic" in lowered:
                entry_type = "dynamic"
            elif "static" in lowered or "estatic" in lowered:
                entry_type = "static"

            # VLAN: prioriza números pequenos próximos ao MAC e ignora octetos/partes do MAC.
            scrubbed = line[:match.start()] + " " * (match.end() - match.start()) + line[match.end():]
            vlan = None
            for token in VLAN_RE.findall(scrubbed):
                number = int(token)
                if 1 <= number <= 4094:
                    vlan = number
                    break

            key = (mac, port, vlan)
            if key not in seen:
                entries.append(MacEntry(mac=mac, port=port, vlan=vlan, entry_type=entry_type))
                seen.add(key)
        return entries


class AclPlanner:
    def __init__(self, profile: SwitchProfile) -> None:
        self.profile = profile

    def build_prefix_allowlist(
        self,
        prefixes: Iterable[str],
        ports: Iterable[str],
        *,
        acl_id: int = 2001,
    ) -> list[str]:
        if not self.profile.prefix_acl_supported:
            raise ValueError(
                f"O perfil {self.profile.display_name} não possui geração automática de ACL por OUI habilitada. "
                "Use auditoria apenas ou cadastre MACs exatos até validar a sintaxe específica do firmware."
            )
        if self.profile.key != "intelbras_serie3000":
            raise ValueError("Gerador de prefixo implementado apenas para Intelbras Série 3000.")
        if not 2001 <= acl_id <= 3000:
            raise ValueError("Na Série 3000, use ACL MAC padrão no intervalo 2001-3000.")

        normalized = sorted({normalize_prefix(p) for p in prefixes})
        normalized_ports = [p.strip() for p in ports if p.strip()]
        if not normalized:
            raise ValueError("Nenhum OUI informado; recuso gerar uma política deny-by-default vazia.")
        if not normalized_ports:
            raise ValueError("Nenhuma porta de acesso informada.")

        commands = ["configure terminal", f"mac access-list standard {acl_id}"]
        seq = 10
        for prefix in normalized:
            base, wildcard = prefix_to_serie3000_acl(prefix)
            commands.append(f"{seq} permit {base} {wildcard}")
            seq += 10
        # A Série 3000 documenta descarte padrão quando nenhuma regra corresponde.
        commands.append("exit")
        for port in normalized_ports:
            commands.extend([
                f"interface {port}",
                f"mac access-group {acl_id} in",
                "exit",
            ])
        commands.append("end")
        return commands

    def build_exact_allowlist(
        self,
        macs: Iterable[str],
        ports: Iterable[str],
        *,
        acl_id: int,
    ) -> list[str]:
        normalized_macs = sorted({normalize_mac(m) for m in macs})
        normalized_ports = [p.strip() for p in ports if p.strip()]
        if not normalized_macs or not normalized_ports:
            raise ValueError("MACs e portas são obrigatórios.")

        if self.profile.key == "intelbras_s2050g_a":
            commands = ["configure terminal", f"mac access-list {acl_id}"]
            rule = 1
            for mac in normalized_macs:
                source = mac_to_intelbras(mac)
                commands.append(f"rule {rule} permit {source} 0000.0000.0000 any")
                rule += 1
            commands.append(f"rule {rule} deny any any")
            commands.append("exit")
            for port in normalized_ports:
                commands.extend([f"interface {port}", f"mac-access-group {acl_id} in", "exit"])
            commands.append("end")
            return commands

        if self.profile.key == "intelbras_s_series":
            commands = ["configure terminal", f"mac access-list extended {acl_id}"]
            rule = 0
            for mac in normalized_macs:
                commands.append(f"{rule} permit host {mac_to_intelbras(mac)} any")
                rule += 1
            if rule > 9:
                raise ValueError("S-Series suporta somente 10 regras por lista MAC ACL; reduza a lista ou divida a política.")
            commands.append(f"{rule} deny any any")
            commands.append("exit")
            for port in normalized_ports:
                commands.extend([f"interface {port}", f"mac access-list {acl_id} commit", "exit"])
            commands.append("end")
            return commands

        raise ValueError("Allowlist exata ainda não implementada para este perfil.")


class SwitchSSHClient:
    """Cliente não interativo usando OpenSSH do Windows; requer autenticação por chave."""

    def __init__(self, *, ssh_path: str | None = None, timeout: int = 20) -> None:
        self.ssh_path = ssh_path or shutil.which("ssh.exe") or shutil.which("ssh")
        self.timeout = timeout

    def run(
        self,
        host: str,
        username: str,
        commands: Iterable[str],
        *,
        identity_file: str | None = None,
        port: int = 22,
    ) -> subprocess.CompletedProcess[str]:
        if not self.ssh_path:
            raise RuntimeError("OpenSSH client não encontrado no Windows.")
        cmd = [
            self.ssh_path,
            "-T",
            "-p", str(port),
            "-o", "BatchMode=yes",
            "-o", "ConnectTimeout=8",
        ]
        if identity_file:
            cmd.extend(["-i", identity_file])
        cmd.append(f"{username}@{host}")
        payload = "\n".join(commands) + "\n"
        return subprocess.run(
            cmd,
            input=payload,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=self.timeout,
            shell=False,
        )


class NetworkNACModule:
    def __init__(self, allowlist_path: Path) -> None:
        self.allowlist_path = allowlist_path
        self.allowlist = OuiAllowlist(allowlist_path)

    def parse_and_classify(self, mac_table_text: str) -> list[Classification]:
        return [self.allowlist.classify(entry) for entry in MacTableParser.parse(mac_table_text)]

    def audit_summary(self, mac_table_text: str) -> dict:
        classified = self.parse_and_classify(mac_table_text)
        unauthorized = [item for item in classified if not item.authorized]
        by_port: dict[str, dict[str, int]] = {}
        for item in classified:
            port = item.port or "desconhecida"
            stats = by_port.setdefault(port, {"authorized": 0, "unauthorized": 0})
            stats["authorized" if item.authorized else "unauthorized"] += 1
        return {
            "total": len(classified),
            "authorized": len(classified) - len(unauthorized),
            "unauthorized": len(unauthorized),
            "by_port": by_port,
            "items": [asdict(item) for item in classified],
        }

    def plan_prefix_acl(self, profile_key: str, ports: Iterable[str], *, acl_id: int = 2001) -> list[str]:
        profile = PROFILES[profile_key]
        planner = AclPlanner(profile)
        return planner.build_prefix_allowlist(self.allowlist.allowed_prefixes, ports, acl_id=acl_id)
