from __future__ import annotations

from core.executor import RemoteExecutor
from core.progress import ProgressParser
from core.result import CommandResult


class RepairModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def sfc_scan(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(
            host,
            "sfc /scannow",
            timeout=1800,
        )

    def dism_checkhealth(self, host: str) -> CommandResult:
        return self.executor.execute_cmd(
            host,
            "DISM /Online /Cleanup-Image /CheckHealth",
            timeout=600,
        )

    def dism_scanhealth(self, host: str) -> CommandResult:
        return self.executor.execute_cmd(
            host,
            "DISM /Online /Cleanup-Image /ScanHealth",
            timeout=1800,
        )

    def dism_restorehealth(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(
            host,
            "DISM /Online /Cleanup-Image /RestoreHealth",
            timeout=3600,
        )

    def dism_restorehealth_progress(
        self,
        host: str,
        on_progress=None,
    ) -> CommandResult:
        def on_line(text: str) -> None:
            event = ProgressParser.parse(text)
            if on_progress:
                on_progress(event)

        return self.executor.execute_cmd_stream(
            host,
            "DISM /Online /Cleanup-Image /RestoreHealth",
            timeout=3600,
            on_line=on_line,
        )

    def component_store(self, host: str) -> CommandResult:
        return self.executor.execute_cmd(
            host,
            "DISM /Online /Cleanup-Image /AnalyzeComponentStore",
            timeout=900,
        )

    def component_cleanup(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(
            host,
            "DISM /Online /Cleanup-Image /StartComponentCleanup",
            timeout=1800,
        )

    def chkdsk_scan(self, host: str) -> CommandResult:
        return self.executor.execute_cmd(host, "chkdsk C: /scan", timeout=1800)

    def reset_store(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_powershell(
            host,
            "Start-Process wsreset.exe -Wait -ErrorAction Stop",
            timeout=300,
        )

    def repository_consistency(self, host: str) -> CommandResult:
        return self.executor.execute_cmd(
            host,
            "winmgmt /verifyrepository",
            timeout=300,
        )
