from __future__ import annotations

import socket
from concurrent.futures import ThreadPoolExecutor
from typing import Any

from core.host_identity import HostIdentity
from core.result import CommandResult


class ConnectivityDiagnostics:
    """Diagnóstico em camadas: DNS/rede → portas → autenticação → transporte."""

    def __init__(self, executor) -> None:
        self.executor = executor

    @staticmethod
    def _tcp(host: str, port: int, timeout: float = 1.2) -> bool:
        try:
            with socket.create_connection((host, port), timeout=timeout):
                return True
        except OSError:
            return False

    @staticmethod
    def _failure(
        host: str,
        action: str,
        message: str,
    ) -> CommandResult:
        return CommandResult.failure(host, action, message)

    @staticmethod
    def _auth_failure(result: CommandResult | None) -> bool:
        if not result or result.success:
            return False
        text = f"{result.stderr}\n{result.stdout}".casefold()
        markers = (
            "access is denied",
            "acesso negado",
            "logon failure",
            "falha de logon",
            "cannotuseipaddress",
            "trustedhosts",
            "authentication",
            "autenticação",
        )
        return any(marker in text for marker in markers)

    def run(self, host: str) -> dict[str, Any]:
        desc = HostIdentity.describe(host)
        if desc.is_local:
            return {
                "host": host,
                "is_local": True,
                "addresses": list(desc.resolved_addresses),
                "dns": True,
                "ping": True,
                "tcp_445": None,
                "tcp_5985": None,
                "tcp_5986": None,
                "winrm": None,
                "admin_share": None,
                "psexec_available": bool(self.executor.psexec_path),
                "psexec": None,
                "selected_transport": "local",
                "state": "READY_LOCAL",
                "diagnosis": (
                    "Alvo local detectado; WinRM, ADMIN$ e PsExec foram ignorados."
                ),
            }

        addresses = list(desc.resolved_addresses)
        dns_ok = bool(addresses)

        with ThreadPoolExecutor(
            max_workers=4,
            thread_name_prefix="central-n2-preflight",
        ) as pool:
            ping_future = pool.submit(self.executor.ping, host)
            tcp445_future = pool.submit(self._tcp, host, 445)
            tcp5985_future = pool.submit(self._tcp, host, 5985)
            tcp5986_future = pool.submit(self._tcp, host, 5986)

            ping = ping_future.result()
            tcp445 = tcp445_future.result()
            tcp5985 = tcp5985_future.result()
            tcp5986 = tcp5986_future.result()

        if tcp5985:
            winrm = self.executor.test_winrm(host)
        else:
            winrm = self._failure(
                host,
                "test_winrm",
                "TCP 5985 indisponível; Test-WSMan não foi executado.",
            )

        if tcp445:
            admin = self.executor.test_admin_share(host)
        else:
            admin = self._failure(
                host,
                "admin_share",
                "TCP 445 indisponível; ADMIN$ não foi testado.",
            )

        psexec: CommandResult | None = None
        if (
            admin.success
            and bool(self.executor.psexec_path)
        ):
            psexec = self.executor.test_psexec(host)
        elif not self.executor.psexec_path:
            psexec = self._failure(
                host,
                "psexec",
                "PsExec não está disponível na estação administrativa.",
            )
        else:
            psexec = self._failure(
                host,
                "psexec",
                "ADMIN$ indisponível; PsExec não foi testado.",
            )

        selected = self.executor.select_transport(
            host,
            refresh=True,
            winrm_result=winrm,
            psexec_result=psexec,
        )

        if not dns_ok:
            state = "DNS_FAILED"
            diagnosis = "Falha de resolução DNS."
        elif winrm.success:
            state = "READY_WINRM"
            diagnosis = "WinRM autenticado e utilizável."
        elif psexec.success:
            state = "READY_PSEXEC"
            diagnosis = "WinRM indisponível; PsExec/ADMIN$ validado."
        elif self._auth_failure(winrm) or self._auth_failure(admin) or self._auth_failure(psexec):
            state = "AUTHENTICATION_FAILED"
            diagnosis = (
                "A rede responde, mas a autenticação/autorização administrativa falhou."
            )
        elif not ping.success and not tcp445 and not tcp5985 and not tcp5986:
            state = "NETWORK_UNREACHABLE"
            diagnosis = (
                "Nenhum caminho administrativo conhecido respondeu; "
                "verifique rede, firewall e rota."
            )
        else:
            state = "NO_USABLE_TRANSPORT"
            diagnosis = (
                "A estação está alcançável, mas nenhum transporte administrativo "
                "foi validado."
            )

        return {
            "host": host,
            "is_local": False,
            "addresses": addresses,
            "dns": dns_ok,
            "ping": ping.success,
            "tcp_445": tcp445,
            "tcp_5985": tcp5985,
            "tcp_5986": tcp5986,
            "winrm": winrm.success,
            "admin_share": admin.success,
            "psexec_available": bool(self.executor.psexec_path),
            "psexec": psexec.success,
            "selected_transport": selected,
            "state": state,
            "diagnosis": diagnosis,
        }
