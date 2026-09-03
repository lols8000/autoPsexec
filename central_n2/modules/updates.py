from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class UpdatesModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def status(self, host: str) -> CommandResult:
        script = r'''
$history = Get-HotFix -ErrorAction SilentlyContinue | Sort-Object InstalledOn -Descending | Select-Object -First 20 HotFixID,Description,InstalledOn,InstalledBy
$wu = $null
try {
  $session = New-Object -ComObject Microsoft.Update.Session
  $searcher = $session.CreateUpdateSearcher()
  $result = $searcher.Search("IsInstalled=0 and IsHidden=0")
  $pending = @($result.Updates | ForEach-Object { [pscustomobject]@{Title=$_.Title;KB=($_.KBArticleIDs -join ',');RebootRequired=$_.RebootRequired} })
  $wu = [pscustomobject]@{PendingCount=$result.Updates.Count;Pending=$pending}
} catch {
  $wu = [pscustomobject]@{PendingCount=$null;Pending=@();Error=$_.Exception.Message}
}
[pscustomobject]@{History=$history;WindowsUpdate=$wu}
'''
        return self.executor.execute_powershell_json(host, script, timeout=180)

    def trigger_scan(self, host: str) -> CommandResult:
        script = r'''
$uso = Join-Path $env:SystemRoot 'System32\UsoClient.exe'
if(Test-Path $uso){ Start-Process $uso -ArgumentList 'StartScan' -WindowStyle Hidden; 'Busca iniciada.' }
else { throw 'UsoClient.exe não encontrado.' }
'''
        return self.executor.execute_remote_powershell_with_fallback(host, script, timeout=60)

    def reset_components(self, host: str) -> CommandResult:
        script = r'''
$services = 'bits','wuauserv','cryptsvc'
foreach($s in $services){ Stop-Service $s -Force -ErrorAction SilentlyContinue }
$stamp = Get-Date -Format yyyyMMddHHmmss
if(Test-Path "$env:SystemRoot\SoftwareDistribution") { Rename-Item "$env:SystemRoot\SoftwareDistribution" "SoftwareDistribution.$stamp.bak" -ErrorAction Stop }
if(Test-Path "$env:SystemRoot\System32\catroot2") { Rename-Item "$env:SystemRoot\System32\catroot2" "catroot2.$stamp.bak" -ErrorAction Stop }
foreach($s in $services){ Start-Service $s -ErrorAction SilentlyContinue }
'Reset dos componentes concluído.'
'''
        return self.executor.execute_remote_powershell_with_fallback(host, script, timeout=180)
