from __future__ import annotations

import base64
import json
from typing import Any

from core.result import CommandResult
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

    def _run_ps(
        self,
        host: str,
        script: str,
        *,
        action: str,
        timeout: int | None = None,
    ) -> CommandResult:
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

    def test(self, host: str) -> CommandResult:
        """Valida listener, autenticação e execução PowerShell remota."""
        safe = self._safe(host)
        marker = "CENTRAL_N2_WINRM_OK"
        return self._run_ps(
            host,
            (
                "$ErrorActionPreference='Stop'; "
                f"Test-WSMan -ComputerName '{safe}' | Out-Null; "
                f"$r = Invoke-Command -ComputerName '{safe}' "
                f"-ScriptBlock {{ '{marker}' }}; "
                f"if ($r -ne '{marker}') {{ "
                "throw 'WinRM respondeu, mas Invoke-Command não foi validado.' "
                "}; "
                f"'{marker}'"
            ),
            action="test_winrm",
            timeout=20,
        )

    def execute_powershell(
        self,
        host: str,
        script: str,
        *,
        timeout: int | None = None,
    ) -> CommandResult:
        safe = self._safe(host)
        wrapped = (
            "$ErrorActionPreference='Stop'; "
            f"Invoke-Command -ComputerName '{safe}' -ScriptBlock {{ "
            "$ErrorActionPreference='Stop'; "
            f"{script} "
            "}"
        )
        return self._run_ps(
            host,
            wrapped,
            action="powershell_remote",
            timeout=timeout,
        )

    def execute_cmd(
        self,
        host: str,
        command: str,
        *,
        timeout: int | None = None,
    ) -> CommandResult:
        """Executa CMD remoto e propaga o exit code do processo remoto.

        O exit code do powershell.exe local não é suficiente para dizer se o
        cmd.exe remoto teve sucesso. Por isso o destino devolve um envelope JSON.
        """
        safe_host = self._safe(host)
        encoded_command = base64.b64encode(
            command.encode("utf-16le")
        ).decode("ascii")
        script = f"""
$ErrorActionPreference='Stop'
$__command=[Text.Encoding]::Unicode.GetString(
    [Convert]::FromBase64String('{encoded_command}')
)
$__remote=Invoke-Command -ComputerName '{safe_host}' -ScriptBlock {{
    param([string]$Command)
    $output = & cmd.exe /d /c $Command 2>&1 | Out-String
    [pscustomobject]@{{
        ExitCode = [int]$LASTEXITCODE
        Output = $output
    }}
}} -ArgumentList $__command
$__remote | ConvertTo-Json -Depth 4 -Compress
"""
        wrapper = self._run_ps(
            host,
            script,
            action="cmd_remote",
            timeout=timeout,
        )
        if not wrapper.success:
            wrapper.command = command
            return wrapper

        try:
            payload: Any = json.loads(wrapper.stdout.strip())
            if not isinstance(payload, dict) or "ExitCode" not in payload:
                raise ValueError("envelope WinRM inválido")
            exit_code = int(payload["ExitCode"])
            output = str(payload.get("Output") or "").strip()
        except (ValueError, TypeError, json.JSONDecodeError) as exc:
            result = CommandResult.failure(
                host,
                command,
                "Não foi possível confirmar o exit code do comando remoto via WinRM.",
                return_code=125,
                transport=self.name,
            )
            result.duration_ms = wrapper.duration_ms
            result.metadata["wrapper_stdout"] = wrapper.stdout
            result.metadata["parse_error"] = str(exc)
            return result.mark_indeterminate(
                "O wrapper WinRM terminou, mas o resultado remoto não pôde ser validado."
            )

        return CommandResult(
            success=exit_code == 0,
            command=command,
            host=host,
            stdout=output,
            stderr="" if exit_code == 0 else f"Comando remoto retornou exit code {exit_code}.",
            return_code=exit_code,
            duration_ms=wrapper.duration_ms,
            transport=self.name,
            metadata={"remote_exit_code_confirmed": True},
        )
