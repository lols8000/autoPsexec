from __future__ import annotations

import base64

from core.result import CommandResult


class RemoteProcessController:
    """Controla processos PowerShell remotos explicitamente rastreados."""

    def __init__(self, executor) -> None:
        self.executor = executor

    def start_powershell(
        self,
        host: str,
        script: str,
    ) -> CommandResult:
        encoded = base64.b64encode(
            script.encode("utf-16le")
        ).decode("ascii")
        payload = (
            "$p=Start-Process powershell.exe "
            "-ArgumentList '-NoProfile','-EncodedCommand',"
            f"'{encoded}' "
            "-PassThru -WindowStyle Hidden;"
            "[pscustomobject]@{PID=$p.Id}"
        )
        return self.executor.execute_mutating_powershell_json(
            host,
            payload,
            timeout=60,
        )

    def cancel(
        self,
        host: str,
        pid: int,
    ) -> CommandResult:
        process_id = int(pid)
        if process_id <= 0:
            return CommandResult.failure(
                host,
                "cancel_remote_process",
                "PID inválido.",
            )
        return self.executor.execute_mutating_powershell(
            host,
            (
                f"Stop-Process -Id {process_id} "
                "-Force -ErrorAction Stop"
            ),
            timeout=60,
        )
