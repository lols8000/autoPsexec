from __future__ import annotations

from pathlib import Path
from typing import Any

from core.config import ConfigLoader
from core.executor import RemoteExecutor
from core.result import CommandResult


class SoftwareModule:
    def __init__(
        self,
        executor: RemoteExecutor,
        settings_path: str | Path,
        *,
        settings: dict[str, Any] | None = None,
    ) -> None:
        self.executor = executor
        self.settings = settings or ConfigLoader(settings_path).settings

    @property
    def catalog(self) -> dict[str, dict[str, Any]]:
        return self.settings.get("software", {})

    def list_installed(self, host: str) -> CommandResult:
        script = """
$paths = @(
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
Get-ItemProperty $paths -ErrorAction SilentlyContinue |
    Where-Object DisplayName |
    Select-Object DisplayName,DisplayVersion,Publisher,InstallDate |
    Sort-Object DisplayName -Unique
"""
        return self.executor.execute_powershell_json(host, script, timeout=120)

    def winget_available(self, host: str) -> CommandResult:
        script = """
$cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
[pscustomobject]@{
    Available = [bool]$cmd
    Path = if ($cmd) { $cmd.Source } else { $null }
}
"""
        return self.executor.execute_powershell_json(host, script)

    def _item(
        self,
        host: str,
        key: str,
    ) -> dict[str, Any] | CommandResult:
        item = self.catalog.get(key)
        if item:
            return item
        return CommandResult.failure(
            host,
            key,
            f"Software '{key}' não existe no catálogo.",
        )

    def _winget_id(
        self,
        host: str,
        key: str,
    ) -> str | CommandResult:
        item = self._item(host, key)
        if isinstance(item, CommandResult):
            return item
        winget_id = str(item.get("winget_id") or "").strip()
        if not winget_id:
            return CommandResult.failure(
                host,
                key,
                f"Software '{key}' não possui winget_id.",
            )
        return winget_id.replace("'", "''")

    def install_catalog_item(self, host: str, key: str) -> CommandResult:
        winget_id = self._winget_id(host, key)
        if isinstance(winget_id, CommandResult):
            return winget_id
        return self.executor.execute_mutating_powershell(
            host,
            (
                f"winget install --id '{winget_id}' --exact --silent "
                "--accept-package-agreements --accept-source-agreements "
                "--disable-interactivity"
            ),
            timeout=600,
        )

    def upgrade_catalog_item(self, host: str, key: str) -> CommandResult:
        winget_id = self._winget_id(host, key)
        if isinstance(winget_id, CommandResult):
            return winget_id
        return self.executor.execute_mutating_powershell(
            host,
            (
                f"winget upgrade --id '{winget_id}' --exact --silent "
                "--accept-package-agreements --accept-source-agreements "
                "--disable-interactivity"
            ),
            timeout=600,
        )

    def uninstall_catalog_item(self, host: str, key: str) -> CommandResult:
        winget_id = self._winget_id(host, key)
        if isinstance(winget_id, CommandResult):
            return winget_id
        return self.executor.execute_mutating_powershell(
            host,
            (
                f"winget uninstall --id '{winget_id}' --exact --silent "
                "--disable-interactivity"
            ),
            timeout=600,
        )
