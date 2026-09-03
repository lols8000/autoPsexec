from __future__ import annotations

import base64
from pathlib import Path

from core.result import CommandResult
from .base import RunLocal, Transport, Utf8Prefix


class PsExecTransport(Transport):
    name = "psexec"

    def __init__(self, runner: RunLocal, utf8_prefix: Utf8Prefix, executable: str | None) -> None:
        self._runner = runner
        self._utf8_prefix = utf8_prefix
        self.executable = executable

    def available(self) -> bool:
        return bool(self.executable and Path(self.executable).exists())

    def _base(self, host: str, *, system: bool = False) -> list[str]:
        if not self.available():
            return []
        cmd = [str(self.executable), "-accepteula", "-nobanner", f"\\\\{host}"]
        if system:
            cmd.append("-s")
        return cmd

    def execute_raw(
        self,
        host: str,
        executable: str,
        args: list[str] | None = None,
        *,
        system: bool = False,
        timeout: int | None = None,
        output_encoding: str | None = None,
    ) -> CommandResult:
        if not self.available():
            return CommandResult.failure(
                host, executable,
                "PsExec não encontrado. Configure psexec_path ou instale Sysinternals.",
                transport=self.name,
            )
        cmd = self._base(host, system=system) + [executable, *(args or [])]
        result = self._runner(
            cmd, host=host, action="psexec", timeout=timeout,
            output_encoding=output_encoding,
        )
        result.transport = self.name
        return result

    def test(self, host: str) -> CommandResult:
        return self.execute_raw(host, "cmd.exe", ["/d", "/c", "echo CENTRAL_N2_OK"], timeout=15)

    def execute_powershell(self, host: str, script: str, *, timeout: int | None = None) -> CommandResult:
        payload = self._utf8_prefix() + "$ErrorActionPreference='Stop'; " + script
        encoded = base64.b64encode(payload.encode("utf-16le")).decode("ascii")
        return self.execute_raw(
            host,
            "powershell.exe",
            ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
            timeout=timeout,
            output_encoding="utf-8",
        )

    def execute_cmd(self, host: str, command: str, *, timeout: int | None = None) -> CommandResult:
        return self.execute_raw(host, "cmd.exe", ["/d", "/c", command], timeout=timeout)
