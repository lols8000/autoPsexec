from __future__ import annotations

import json
from pathlib import Path
from typing import Dict, Any

from core.executor import RemoteExecutor
from core.result import CommandResult


class SoftwareModule:
    def __init__(self, executor: RemoteExecutor, settings_path: str | Path) -> None:
        self.executor = executor
        self.settings_path = Path(settings_path)
        self.settings = json.loads(self.settings_path.read_text(encoding="utf-8"))

    @property
    def catalog(self) -> Dict[str, Dict[str, Any]]:
        return self.settings.get("software", {})

    def list_installed(self, host: str) -> CommandResult:
        script = r'''
$paths = @(
 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
Get-ItemProperty $paths -ErrorAction SilentlyContinue |
 Where-Object DisplayName |
 Select-Object DisplayName,DisplayVersion,Publisher,InstallDate |
 Sort-Object DisplayName -Unique
'''
        return self.executor.execute_powershell_json(host, script, timeout=120)

    def winget_available(self, host: str) -> CommandResult:
        script = r'''
$cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
[pscustomobject]@{ Available = [bool]$cmd; Path = if($cmd){$cmd.Source}else{$null} }
'''
        return self.executor.execute_powershell_json(host, script)

    def install_catalog_item(self, host: str, key: str) -> CommandResult:
        item = self.catalog.get(key)
        if not item:
            return CommandResult.failure(host, key, f"Software '{key}' não existe no catálogo.")
        winget_id = item["winget_id"].replace("'", "''")
        script = (
            f"winget install --id '{winget_id}' --exact --silent "
            "--accept-package-agreements --accept-source-agreements --disable-interactivity"
        )
        return self.executor.execute_remote_powershell_with_fallback(host, script, timeout=600)

    def upgrade_catalog_item(self, host: str, key: str) -> CommandResult:
        item = self.catalog.get(key)
        if not item:
            return CommandResult.failure(host, key, f"Software '{key}' não existe no catálogo.")
        winget_id = item["winget_id"].replace("'", "''")
        script = (
            f"winget upgrade --id '{winget_id}' --exact --silent "
            "--accept-package-agreements --accept-source-agreements --disable-interactivity"
        )
        return self.executor.execute_remote_powershell_with_fallback(host, script, timeout=600)

    def uninstall_catalog_item(self, host: str, key: str) -> CommandResult:
        item = self.catalog.get(key)
        if not item:
            return CommandResult.failure(host, key, f"Software '{key}' não existe no catálogo.")
        winget_id = item["winget_id"].replace("'", "''")
        return self.executor.execute_remote_powershell_with_fallback(
            host,
            f"winget uninstall --id '{winget_id}' --exact --silent --disable-interactivity",
            timeout=600,
        )
