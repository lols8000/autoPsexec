from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class StartupModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def overview(self, host: str) -> CommandResult:
        script = r'''
$run=@()
$run += Get-ItemProperty 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run' -ErrorAction SilentlyContinue | Select-Object * -ExcludeProperty PS*
$run += Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -ErrorAction SilentlyContinue | Select-Object * -ExcludeProperty PS*
$startup=Get-CimInstance Win32_StartupCommand -ErrorAction SilentlyContinue | Select-Object Name,Command,Location,User
$tasks=Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {$_.State -ne 'Disabled'} | Select-Object TaskName,TaskPath,State
$autoStopped=Get-CimInstance Win32_Service | Where-Object {$_.StartMode -eq 'Auto' -and $_.State -ne 'Running'} | Select-Object Name,DisplayName,State,StartMode
[pscustomobject]@{Run=$run;Startup=$startup;ScheduledTasks=$tasks;AutomaticStoppedServices=$autoStopped}
'''
        return self.executor.execute_powershell_json(host, script, timeout=180)
