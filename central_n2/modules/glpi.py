from __future__ import annotations

import shutil
from pathlib import Path, PureWindowsPath
from typing import Any

from core.config import ConfigLoader
from core.executor import RemoteExecutor
from core.result import CommandResult


class GLPIModule:
    def __init__(
        self,
        executor: RemoteExecutor,
        settings_path: str | Path,
        *,
        settings: dict[str, Any] | None = None,
    ) -> None:
        self.executor = executor
        self.settings = settings or ConfigLoader(settings_path).settings
        self.config = self.settings.get("glpi", {})

    def status(self, host: str) -> CommandResult:
        names = self.config.get("service_names", ["glpi-agent", "GLPI-Agent"])
        ps_names = ",".join(
            f"'{str(name).replace(chr(39), chr(39) * 2)}'"
            for name in names
        )
        script = fr"""
$services = Get-Service -ErrorAction SilentlyContinue |
    Where-Object {{ $_.Name -in @({ps_names}) -or $_.DisplayName -match 'GLPI' }}
$exe = Get-ChildItem 'C:\Program Files\GLPI-Agent' -Filter glpi-agent.exe -Recurse -ErrorAction SilentlyContinue |
    Select-Object -First 1
[pscustomobject]@{{
    Installed = [bool]($services -or $exe)
    Running = [bool](@($services | Where-Object Status -eq 'Running').Count -gt 0)
    Services = @($services | Select-Object Name,DisplayName,Status,StartType)
    Executable = if ($exe) {{ $exe.FullName }} else {{ $null }}
}}
"""
        return self.executor.execute_powershell_json(host, script)

    def copy_installer(self, host: str) -> CommandResult:
        source = self.config.get("installer_source")
        remote = str(
            self.config.get(
                "remote_installer_path",
                r"C:\glpiagentinstall.vbs",
            )
        )
        if not source:
            return CommandResult.failure(
                host,
                "copy_glpi_installer",
                "installer_source não configurado.",
            )

        remote_path = PureWindowsPath(remote)
        if not remote_path.drive:
            return CommandResult.failure(
                host,
                "copy_glpi_installer",
                "remote_installer_path deve ser um caminho absoluto do Windows.",
            )

        drive = remote_path.drive.rstrip(":")
        relative_parts = remote_path.parts[1:]
        destination = (
            f"\\\\{host}\\{drive}$\\"
            + "\\".join(relative_parts)
        )

        parent = str(PureWindowsPath(destination).parent)
        try:
            Path(parent).mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)
        except OSError as exc:
            return CommandResult.failure(
                host,
                "copy_glpi_installer",
                str(exc),
            )

        result = CommandResult(
            True,
            "copy_glpi_installer",
            host,
            stdout=f"Copiado para {destination}",
        )
        if self.executor.logger:
            self.executor.logger.log_result("copy_glpi_installer", result)
        return result

    def install_or_repair(self, host: str) -> CommandResult:
        copied = self.copy_installer(host)
        if not copied.success:
            return copied

        remote = str(
            self.config.get(
                "remote_installer_path",
                r"C:\glpiagentinstall.vbs",
            )
        )
        result = self.executor.execute_mutating_cmd(
            host,
            f'cscript.exe //nologo "{remote}"',
            timeout=600,
        )
        if result.success:
            self.executor.execute_mutating_cmd(
                host,
                f'del /f /q "{remote}"',
            )
        return result

    def restart_service(self, host: str) -> CommandResult:
        script = r"""
$svc = Get-Service -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match 'glpi' -or $_.DisplayName -match 'GLPI' } |
    Select-Object -First 1
if (-not $svc) { throw 'Serviço GLPI Agent não encontrado.' }
Restart-Service -Name $svc.Name -Force -ErrorAction Stop
Get-Service -Name $svc.Name | Select-Object Name,Status
"""
        return self.executor.execute_mutating_powershell_json(host, script)

    def force_inventory(self, host: str) -> CommandResult:
        script = r"""
$exe = Get-ChildItem 'C:\Program Files\GLPI-Agent' -Filter glpi-agent.exe -Recurse -ErrorAction SilentlyContinue |
    Select-Object -First 1
if (-not $exe) { throw 'glpi-agent.exe não encontrado.' }
& $exe.FullName --force
if ($LASTEXITCODE -ne 0) { throw "GLPI Agent retornou exit code $LASTEXITCODE" }
"""
        return self.executor.execute_mutating_powershell(
            host,
            script,
            timeout=300,
        )

    def recent_log(self, host: str, lines: int = 80) -> CommandResult:
        lines = max(1, min(int(lines), 500))
        script = fr"""
$paths = @(
    'C:\Program Files\GLPI-Agent\logs\glpi-agent.log',
    'C:\Program Files\GLPI-Agent\var\log\glpi-agent.log'
)
$log = $paths | Where-Object {{ Test-Path $_ }} | Select-Object -First 1
if (-not $log) {{ throw 'Log do GLPI Agent não encontrado.' }}
Get-Content $log -Tail {lines}
"""
        return self.executor.execute_powershell(
            host,
            script,
            timeout=120,
        )
