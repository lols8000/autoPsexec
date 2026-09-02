from __future__ import annotations

import json
import shutil
import socket
import subprocess
import time
from pathlib import Path
from typing import Iterable, Optional

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
        self.psexec_path = psexec_path or self._discover_psexec()

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
    def resolve_host(host: str) -> str | None:
        try:
            return socket.gethostbyname(host)
        except OSError:
            return None

    def ping(self, host: str) -> CommandResult:
        cmd = ["ping", "-n", "1", "-w", "1200", host]
        return self._run_local(cmd, host=host, action="ping")

    def test_admin_share(self, host: str) -> CommandResult:
        cmd = ["cmd.exe", "/d", "/c", f"dir \\\\{host}\\admin$ >nul 2>&1"]
        return self._run_local(cmd, host=host, action="admin_share")

    def test_winrm(self, host: str) -> CommandResult:
        script = (
            f"$ErrorActionPreference='Stop'; "
            f"Test-WSMan -ComputerName '{host}' | Out-Null; 'OK'"
        )
        return self._run_powershell_local(script, host=host, action="test_winrm", timeout=15)

    def execute_powershell(self, host: str, script: str, *, timeout: int | None = None) -> CommandResult:
        wrapped = (
            "$ErrorActionPreference='Stop'; "
            f"Invoke-Command -ComputerName '{host}' -ScriptBlock {{ {script} }}"
        )
        result = self._run_powershell_local(
            wrapped,
            host=host,
            action="powershell_remote",
            timeout=timeout,
        )
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
        result = self._run_local(cmd, host=host, action="psexec", timeout=timeout)
        result.transport = "psexec"
        return result

    def execute_cmd(self, host: str, command: str, *, timeout: int | None = None) -> CommandResult:
        winrm = self.test_winrm(host)
        if winrm.success:
            escaped = command.replace("'", "''")
            return self.execute_powershell(
                host,
                f"cmd.exe /d /c '{escaped}'",
                timeout=timeout,
            )
        return self.execute_psexec(host, "cmd.exe", ["/d", "/c", command], timeout=timeout)

    def execute_remote_powershell_with_fallback(
        self,
        host: str,
        script: str,
        *,
        timeout: int | None = None,
    ) -> CommandResult:
        winrm = self.test_winrm(host)
        if winrm.success:
            return self.execute_powershell(host, script, timeout=timeout)

        encoded = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-Command",
                "[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($args[0]))",
                script,
            ],
            capture_output=True,
            text=True,
            timeout=10,
        )
        if encoded.returncode != 0:
            return CommandResult.failure(host, script, encoded.stderr, transport="local")
        return self.execute_psexec(
            host,
            "powershell.exe",
            ["-NoProfile", "-NonInteractive", "-EncodedCommand", encoded.stdout.strip()],
            timeout=timeout,
        )

    def _run_powershell_local(
        self,
        script: str,
        *,
        host: str,
        action: str,
        timeout: int | None = None,
    ) -> CommandResult:
        return self._run_local(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
            host=host,
            action=action,
            timeout=timeout,
        )

    def _run_local(
        self,
        cmd: list[str],
        *,
        host: str,
        action: str,
        timeout: int | None = None,
    ) -> CommandResult:
        started = time.perf_counter()
        printable = subprocess.list2cmdline(cmd)
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout or self.timeout,
                shell=False,
            )
            result = CommandResult(
                success=proc.returncode == 0,
                command=printable,
                host=host,
                stdout=proc.stdout.strip(),
                stderr=proc.stderr.strip(),
                return_code=proc.returncode,
                duration_ms=int((time.perf_counter() - started) * 1000),
            )
        except subprocess.TimeoutExpired:
            result = CommandResult.failure(
                host,
                printable,
                f"Timeout após {timeout or self.timeout}s",
                return_code=124,
            )
            result.duration_ms = int((time.perf_counter() - started) * 1000)
        except OSError as exc:
            result = CommandResult.failure(host, printable, str(exc), return_code=127)
            result.duration_ms = int((time.perf_counter() - started) * 1000)

        if self.logger:
            self.logger.log_result(action, result)
        return result
