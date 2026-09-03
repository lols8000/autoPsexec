from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class CrashesModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def bsod_history(self, host: str, days: int = 30) -> CommandResult:
        days = max(1, min(int(days), 90))
        script = f'''
$start=(Get-Date).AddDays(-{days})
$events=Get-WinEvent -FilterHashtable @{{LogName='System'; Id=1001; StartTime=$start}} -ErrorAction SilentlyContinue |
 Select-Object TimeCreated,Id,ProviderName,Message
$dumps=@()
if(Test-Path 'C:\\Windows\\Minidump'){{$dumps=Get-ChildItem 'C:\\Windows\\Minidump' -Filter *.dmp -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 20 Name,Length,LastWriteTime,FullName}}
$memory=Get-Item 'C:\\Windows\\MEMORY.DMP' -ErrorAction SilentlyContinue | Select-Object Name,Length,LastWriteTime,FullName
[pscustomobject]@{{BugcheckEvents=$events;Minidumps=$dumps;MemoryDump=$memory}}
'''
        return self.executor.execute_powershell_json(host, script, timeout=180)

    def app_crashes(self, host: str, hours: int = 48) -> CommandResult:
        hours = max(1, min(int(hours), 720))
        script = f'''
$start=(Get-Date).AddHours(-{hours})
Get-WinEvent -FilterHashtable @{{LogName='Application'; Id=1000,1001; StartTime=$start}} -ErrorAction SilentlyContinue |
 Select-Object -First 50 TimeCreated,Id,ProviderName,LevelDisplayName,Message
'''
        return self.executor.execute_powershell_json(host, script, timeout=180)
