from __future__ import annotations

import base64
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

from .logger import AuditLogger
from .result import CommandResult


class RemoteExecutor:
    """Executa comandos remotos priorizando PowerShell Remoting e usando PsExec como fallback."""

    def __init__(
        self,
        *,
        psexec_path: str | None = None,
        timeout: int = 60,
        logger: AuditLogger | None = None,
    ) -> None:
        self.timeout = timeout
        self.logger = logger
        configured = Path(psexec_path) if psexec_path else None
        self.psexec_path = str(configured) if configured and configured.exists() else self._discover_psexec()

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
        """Retorna a code page de saída do console Windows, quando disponível."""
        if os.name != "nt":
            return None
        try:
            code_page = int(ctypes.windll.kernel32.GetConsoleOutputCP())
        except (AttributeError, OSError, ValueError):
            return None
        return f"cp{code_page}" if code_page > 0 else None

    @classmethod
    def _decode_output(cls, data: bytes | None, *, preferred: str | None = None) -> str:
        """Decodifica stdout/stderr sem destruir caracteres acentuados.

        PowerShell executado pela Central é explicitamente configurado para UTF-8.
        Para comandos nativos do Windows, a code page do console é priorizada.
        O fallback final usa latin-1 apenas para preservar os bytes em vez de
        substituí-los por U+FFFD (�).
        """
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
        for encoding in (
            preferred,
            "utf-8",
            cls._console_encoding(),
            locale.getpreferredencoding(False),
            "cp850",
            "cp1252",
        ):
            if encoding and encoding.lower() not in {item.lower() for item in candidates}:
                candidates.append(encoding)

        for encoding in candidates:
            try:
                return data.decode(encoding, errors="strict")
            except (LookupError, UnicodeDecodeError):
                continue

        # Último recurso: mapeamento 1:1 para não perder bytes com replacement chars.
        return data.decode("latin-1", errors="strict")

    @staticmethod
    def _powershell_utf8_prefix() -> str:
        # Windows PowerShell 5.1 não garante UTF-8 quando stdout/stderr é redirecionado.
        # Definir Console.OutputEncoding evita mojibake na captura feita pelo Python.
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

    def ping(self, host: str) -> CommandResult:
        return self._run_local(["ping", "-n", "1", "-w", "1200", host], host=host, action="ping")

    def test_admin_share(self, host: str) -> CommandResult:
        return self._run_local(
            ["cmd.exe", "/d", "/c", f"dir \\\\{host}\\admin$ >nul 2>&1"],
            host=host,
            action="admin_share",
        )

    def test_winrm(self, host: str) -> CommandResult:
        safe_host = host.replace("'", "''")
        script = f"$ErrorActionPreference='Stop'; Test-WSMan -ComputerName '{safe_host}' | Out-Null; 'OK'"
        return self._run_powershell_local(script, host=host, action="test_winrm", timeout=15)

    def execute_powershell(self, host: str, script: str, *, timeout: int | None = None) -> CommandResult:
        safe_host = host.replace("'", "''")
        wrapped = f"$ErrorActionPreference='Stop'; Invoke-Command -ComputerName '{safe_host}' -ScriptBlock {{ {script} }}"
        result = self._run_powershell_local(wrapped, host=host, action="powershell_remote", timeout=timeout)
        result.transport = "winrm"
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
        if not self.psexec_path:
            return CommandResult.failure(
                host,
                executable,
                "PsExec não encontrado. Instale Sysinternals PsExec ou configure psexec_path.",
                transport="psexec",
            )
        cmd = [self.psexec_path, "-accepteula", "-nobanner", f"\\\\{host}"]
        if system:
            cmd.append("-s")
        cmd.extend([executable, *list(args)])
        result = self._run_local(
            cmd,
            host=host,
            action="psexec",
            timeout=timeout,
            output_encoding=output_encoding,
        )
        result.transport = "psexec"
        return result

    def execute_cmd(self, host: str, command: str, *, timeout: int | None = None) -> CommandResult:
        if self.test_winrm(host).success:
            escaped = command.replace("'", "''")
            return self.execute_powershell(host, f"cmd.exe /d /c '{escaped}'", timeout=timeout)
        return self.execute_psexec(host, "cmd.exe", ["/d", "/c", command], timeout=timeout)

    def execute_remote_powershell_with_fallback(
        self,
        host: str,
        script: str,
        *,
        timeout: int | None = None,
    ) -> CommandResult:
        if self.test_winrm(host).success:
            return self.execute_powershell(host, script, timeout=timeout)

        remote_script = self._powershell_utf8_prefix() + script
        encoded = base64.b64encode(remote_script.encode("utf-16le")).decode("ascii")
        return self.execute_psexec(
            host,
            "powershell.exe",
            ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
            timeout=timeout,
            output_encoding="utf-8",
        )

    def _run_powershell_local(
        self,
        script: str,
        *,
        host: str,
        action: str,
        timeout: int | None = None,
    ) -> CommandResult:
        utf8_script = self._powershell_utf8_prefix() + script
        return self._run_local(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", utf8_script],
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
