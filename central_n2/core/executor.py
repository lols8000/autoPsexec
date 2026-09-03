from __future__ import annotations

import ctypes
import json
import locale
import os
import shutil
import socket
import subprocess
import time
from pathlib import Path
from typing import Iterable

from .host_identity import HostIdentity
from .logger import AuditLogger
from .result import CommandResult
from .transport import LocalTransport, PsExecTransport, TransportManager, WinRMTransport


class RemoteExecutor:
    """Facade de execução com seleção Local -> WinRM -> PsExec."""

    def __init__(
        self,
        *,
        psexec_path: str | None = None,
        timeout: int = 60,
        logger: AuditLogger | None = None,
        transport_cache_ttl_seconds: float = 120.0,
    ) -> None:
        self.timeout = timeout
        self.logger = logger
        configured = Path(psexec_path) if psexec_path else None
        self.psexec_path = str(configured) if configured and configured.exists() else self._discover_psexec()
        self.local_transport = LocalTransport(self._run_local, self._powershell_utf8_prefix)
        self.winrm_transport = WinRMTransport(self._run_local, self._powershell_utf8_prefix)
        self.psexec_transport = PsExecTransport(self._run_local, self._powershell_utf8_prefix, self.psexec_path)
        self.transport_manager = TransportManager(
            self.local_transport,
            self.winrm_transport,
            self.psexec_transport,
            cache_ttl_seconds=transport_cache_ttl_seconds,
        )

    @staticmethod
    def _discover_psexec() -> str | None:
        candidates = [
            shutil.which("PsExec.exe"),
            shutil.which("psexec.exe"),
            r"C:\Windows\System32\PsExec.exe",
            r"C:\Sysinternals\PsExec.exe",
        ]
        for candidate in candidates:
            if candidate and Path(candidate).exists():
                return str(candidate)
        return None

    @staticmethod
    def _console_encoding() -> str | None:
        if os.name != "nt":
            return None
        try:
            code_page = int(ctypes.windll.kernel32.GetConsoleOutputCP())
        except (AttributeError, OSError, ValueError):
            return None
        return f"cp{code_page}" if code_page > 0 else None

    @classmethod
    def _decode_output(cls, data: bytes | None, *, preferred: str | None = None) -> str:
        if not data:
            return ""
        if data.startswith((b"\xff\xfe", b"\xfe\xff")):
            try:
                return data.decode("utf-16")
            except UnicodeDecodeError:
                pass
        if data.startswith(b"\xef\xbb\xbf"):
            try:
                return data.decode("utf-8-sig")
            except UnicodeDecodeError:
                pass

        candidates: list[str] = []
        for encoding in (preferred, "utf-8", cls._console_encoding(), locale.getpreferredencoding(False), "cp850", "cp1252"):
            if encoding and encoding.lower() not in {item.lower() for item in candidates}:
                candidates.append(encoding)
        for encoding in candidates:
            try:
                return data.decode(encoding, errors="strict")
            except (LookupError, UnicodeDecodeError):
                continue
        return data.decode("latin-1", errors="strict")

    @staticmethod
    def _powershell_utf8_prefix() -> str:
        return (
            "$__centralN2Utf8 = New-Object System.Text.UTF8Encoding($false); "
            "[Console]::OutputEncoding = $__centralN2Utf8; "
            "$OutputEncoding = $__centralN2Utf8; "
        )

    @staticmethod
    def resolve_host(host: str) -> str | None:
        try:
            return socket.gethostbyname(host)
        except OSError:
            return None

    @staticmethod
    def is_local(host: str) -> bool:
        return HostIdentity.is_local(host)

    def select_transport(self, host: str, *, refresh: bool = False) -> str:
        return self.transport_manager.select(host, refresh=refresh).name

    def invalidate_transport(self, host: str) -> None:
        self.transport_manager.invalidate(host)

    def ping(self, host: str) -> CommandResult:
        if self.is_local(host):
            return CommandResult(True, "local", host, stdout="OK", transport="local")
        return self._run_local(["ping", "-n", "1", "-w", "1200", host], host=host, action="ping")

    def test_admin_share(self, host: str) -> CommandResult:
        if self.is_local(host):
            return CommandResult(True, "local", host, stdout="OK", transport="local")
        return self._run_local(
            ["cmd.exe", "/d", "/c", f"dir \\\\{host}\\admin$ >nul 2>&1"],
            host=host,
            action="admin_share",
        )

    def test_winrm(self, host: str) -> CommandResult:
        return self.winrm_transport.test(host)

    @staticmethod
    def _is_transport_failure(result: CommandResult) -> bool:
        text = (result.stderr or "").lower()
        markers = (
            "psremotingtransportexception",
            "cannotconnect",
            "pssessionstatebroken",
            "ws-management",
            "winrm",
            "the client cannot connect",
            "o cliente não conseguiu se conectar",
        )
        return not result.success and any(marker in text for marker in markers)

    def execute_powershell(self, host: str, script: str, *, timeout: int | None = None) -> CommandResult:
        transport = self.transport_manager.select(host)
        result = transport.execute_powershell(host, script, timeout=timeout)
        if (
            transport.name == "winrm"
            and self._is_transport_failure(result)
            and self.psexec_transport.available()
        ):
            self.transport_manager.invalidate(host)
            fallback = self.psexec_transport.execute_powershell(host, script, timeout=timeout)
            fallback.metadata["fallback_from"] = "winrm"
            return fallback
        return result

    def execute_powershell_json(self, host: str, script: str, *, timeout: int | None = None) -> CommandResult:
        result = self.execute_powershell(
            host,
            f"$r = & {{ {script} }}; $r | ConvertTo-Json -Depth 8 -Compress",
            timeout=timeout,
        )
        if result.success and result.stdout.strip():
            try:
                result.data = json.loads(result.stdout.strip())
            except json.JSONDecodeError:
                result.metadata["json_parse_error"] = True
        return result

    def execute_psexec(
        self,
        host: str,
        executable: str,
        args: Iterable[str] = (),
        *,
        system: bool = False,
        timeout: int | None = None,
        output_encoding: str | None = None,
    ) -> CommandResult:
        return self.psexec_transport.execute_raw(
            host,
            executable,
            list(args),
            system=system,
            timeout=timeout,
            output_encoding=output_encoding,
        )

    def execute_cmd(self, host: str, command: str, *, timeout: int | None = None) -> CommandResult:
        transport = self.transport_manager.select(host)
        result = transport.execute_cmd(host, command, timeout=timeout)
        if transport.name == "winrm" and self._is_transport_failure(result) and self.psexec_transport.available():
            self.transport_manager.invalidate(host)
            fallback = self.psexec_transport.execute_cmd(host, command, timeout=timeout)
            fallback.metadata["fallback_from"] = "winrm"
            return fallback
        return result

    def execute_remote_powershell_with_fallback(
        self,
        host: str,
        script: str,
        *,
        timeout: int | None = None,
    ) -> CommandResult:
        return self.execute_powershell(host, script, timeout=timeout)

    def _run_powershell_local(
        self,
        script: str,
        *,
        host: str,
        action: str,
        timeout: int | None = None,
    ) -> CommandResult:
        payload = self._powershell_utf8_prefix() + script
        return self._run_local(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", payload],
            host=host,
            action=action,
            timeout=timeout,
            output_encoding="utf-8",
        )

    def _run_local(
        self,
        cmd: list[str],
        *,
        host: str,
        action: str,
        timeout: int | None = None,
        output_encoding: str | None = None,
    ) -> CommandResult:
        started = time.perf_counter()
        printable = subprocess.list2cmdline(cmd)
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=False,
                timeout=timeout or self.timeout,
                shell=False,
            )
            stdout = self._decode_output(proc.stdout, preferred=output_encoding).strip()
            stderr = self._decode_output(proc.stderr, preferred=output_encoding).strip()
            result = CommandResult(
                success=proc.returncode == 0,
                command=printable,
                host=host,
                stdout=stdout,
                stderr=stderr,
                return_code=proc.returncode,
                duration_ms=int((time.perf_counter() - started) * 1000),
            )
        except subprocess.TimeoutExpired:
            result = CommandResult.failure(host, printable, f"Timeout após {timeout or self.timeout}s", return_code=124)
            result.duration_ms = int((time.perf_counter() - started) * 1000)
        except OSError as exc:
            result = CommandResult.failure(host, printable, str(exc), return_code=127)
            result.duration_ms = int((time.perf_counter() - started) * 1000)

        if self.logger:
            self.logger.log_result(action, result)
        return result
