from __future__ import annotations

from .base import RunLocal, Transport, Utf8Prefix


class WinRMTransport(Transport):
    name = "winrm"

    def __init__(self, runner: RunLocal, utf8_prefix: Utf8Prefix) -> None:
        self._runner = runner
        self._utf8_prefix = utf8_prefix

    def available(self) -> bool:
        return True

    @staticmethod
    def _safe(host: str) -> str:
        return host.replace("'", "''")

    def _run_ps(self, host: str, script: str, *, action: str, timeout: int | None = None):
        payload = self._utf8_prefix() + script
        result = self._runner(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", payload],
            host=host,
            action=action,
            timeout=timeout,
            output_encoding="utf-8",
        )
        result.transport = self.name
        return result

    def test(self, host: str):
        safe = self._safe(host)
        return self._run_ps(
            host,
            f"$ErrorActionPreference='Stop'; Test-WSMan -ComputerName '{safe}' | Out-Null; 'OK'",
            action="test_winrm",
            timeout=15,
        )

    def execute_powershell(self, host: str, script: str, *, timeout: int | None = None):
        safe = self._safe(host)
        wrapped = f"$ErrorActionPreference='Stop'; Invoke-Command -ComputerName '{safe}' -ScriptBlock {{ {script} }}"
        return self._run_ps(host, wrapped, action="powershell_remote", timeout=timeout)

    def execute_cmd(self, host: str, command: str, *, timeout: int | None = None):
        escaped = command.replace("'", "''")
        return self.execute_powershell(host, f"cmd.exe /d /c '{escaped}'", timeout=timeout)
