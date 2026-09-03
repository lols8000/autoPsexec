from __future__ import annotations
from pathlib import Path
from typing import Dict,Any
from core.config import ConfigLoader
from core.executor import RemoteExecutor
from core.result import CommandResult
class SoftwareModule:
    def __init__(self,executor:RemoteExecutor,settings_path:str|Path)->None:self.executor=executor;self.settings=ConfigLoader(settings_path).settings
    @property
    def catalog(self)->Dict[str,Dict[str,Any]]:return self.settings.get("software",{})
    def list_installed(self,host):return self.executor.execute_powershell_json(host,r'''$paths=@('HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*');Get-ItemProperty $paths -ErrorAction SilentlyContinue|Where-Object DisplayName|Select-Object DisplayName,DisplayVersion,Publisher,InstallDate|Sort-Object DisplayName -Unique''',timeout=120)
    def winget_available(self,host):return self.executor.execute_powershell_json(host,r'''$cmd=Get-Command winget.exe -ErrorAction SilentlyContinue;[pscustomobject]@{Available=[bool]$cmd;Path=if($cmd){$cmd.Source}else{$null}}''')
    def _item(self,host,key):
        item=self.catalog.get(key);return item if item else CommandResult.failure(host,key,f"Software '{key}' não existe no catálogo.")
    def install_catalog_item(self,host,key):
        item=self._item(host,key)
        if isinstance(item,CommandResult):return item
        wid=item["winget_id"].replace("'","''");return self.executor.execute_remote_powershell_with_fallback(host,f"winget install --id '{wid}' --exact --silent --accept-package-agreements --accept-source-agreements --disable-interactivity",timeout=600)
    def upgrade_catalog_item(self,host,key):
        item=self._item(host,key)
        if isinstance(item,CommandResult):return item
        wid=item["winget_id"].replace("'","''");return self.executor.execute_remote_powershell_with_fallback(host,f"winget upgrade --id '{wid}' --exact --silent --accept-package-agreements --accept-source-agreements --disable-interactivity",timeout=600)
    def uninstall_catalog_item(self,host,key):
        item=self._item(host,key)
        if isinstance(item,CommandResult):return item
        wid=item["winget_id"].replace("'","''");return self.executor.execute_remote_powershell_with_fallback(host,f"winget uninstall --id '{wid}' --exact --silent --disable-interactivity",timeout=600)
