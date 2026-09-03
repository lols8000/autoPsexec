from __future__ import annotations

import socket
from typing import Any

from core.host_identity import HostIdentity


class ConnectivityDiagnostics:
    def __init__(self, executor) -> None:
        self.executor = executor

    @staticmethod
    def _tcp(host: str, port: int, timeout: float = 1.2) -> bool:
        try:
            with socket.create_connection((host, port), timeout=timeout):
                return True
        except OSError:
            return False

    def run(self, host: str) -> dict[str, Any]:
        desc = HostIdentity.describe(host)
        if desc.is_local:
            return {
                "host": host,
                "is_local": True,
                "addresses": list(desc.resolved_addresses),
                "dns": True,
                "ping": True,
                "tcp_5985": None,
                "tcp_5986": None,
                "winrm": None,
                "admin_share": None,
                "psexec_available": bool(self.executor.psexec_path),
                "selected_transport": "local",
                "diagnosis": "Alvo local detectado; WinRM, ADMIN$ e PsExec foram ignorados.",
            }

        addresses = list(desc.resolved_addresses)
        dns_ok = bool(addresses)
        ping = self.executor.ping(host)
        winrm = self.executor.test_winrm(host)
        admin = self.executor.test_admin_share(host)
        selected = self.executor.select_transport(host, refresh=True, winrm_result=winrm)
        diagnosis = "Conectividade administrativa disponível."
        if not dns_ok:
            diagnosis = "Falha de resolução DNS."
        elif not winrm.success and admin.success and self.executor.psexec_path:
            diagnosis = "WinRM indisponível; fallback PsExec/ADMIN$ disponível."
        elif not winrm.success and not admin.success:
            diagnosis = "WinRM e ADMIN$ indisponíveis; verifique rede, firewall e credenciais."

        return {
            "host": host,
            "is_local": False,
            "addresses": addresses,
            "dns": dns_ok,
            "ping": ping.success,
            "tcp_5985": self._tcp(host, 5985),
            "tcp_5986": self._tcp(host, 5986),
            "winrm": winrm.success,
            "admin_share": admin.success,
            "psexec_available": bool(self.executor.psexec_path),
            "selected_transport": selected,
            "diagnosis": diagnosis,
        }
