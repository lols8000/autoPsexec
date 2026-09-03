from __future__ import annotations

from core.result import CommandResult
from .base import RunLocal, Transport, Utf8Prefix


class LocalTransport(Transport):
    name = "local"

    def __init__(self, runner: RunLocal, utf8_prefix: Utf8Prefix) -> None:
        self._runner = runner
        self._utf8_prefix = utf8_prefix

    def available(self) -> bool:
        return True

    def test(self, host: str) -> CommandResult:
        return CommandResult(True, "local", host, stdout="OK", transport=self.name)

    def execute_powershell(self, host: str, script: str, *, timeout: int | None = None) -> CommandResult:
        payload = self._utf8_prefix() + "$ErrorActionPreference='Stop'; " + script
        result = self._runner(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", payload],
            host=host,
            action="powershell_local",
            timeout=timeout,
            output_encoding="utf-8",
        )
        result.transport = self.name
        return result

    def execute_cmd(self, host: str, command: str, *, timeout: int | None = None) -> CommandResult:
        result = self._runner(
            ["cmd.exe", "/d", "/c", command],
            host=host,
            action="cmd_local",
            timeout=timeout,
        )
        result.transport = self.name
        return result
