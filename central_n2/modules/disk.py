from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class DiskModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def usage(self, host: str) -> CommandResult:
        script = r'''
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
[pscustomobject]@{
 DeviceID='C:'
 SizeGB=[math]::Round($disk.Size/1GB,1)
 FreeGB=[math]::Round($disk.FreeSpace/1GB,1)
 FreePercent=[math]::Round(($disk.FreeSpace/$disk.Size)*100,1)
}
'''
        return self.executor.execute_powershell_json(host, script)

    def top_user_profiles(self, host: str) -> CommandResult:
        script = r'''
Get-ChildItem 'C:\Users' -Directory -Force -ErrorAction SilentlyContinue | ForEach-Object {
  $size = (Get-ChildItem $_.FullName -File -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
  [pscustomobject]@{Profile=$_.Name;Path=$_.FullName;SizeGB=[math]::Round(($size/1GB),2)}
} | Sort-Object SizeGB -Descending
'''
        return self.executor.execute_powershell_json(host, script, timeout=240)

    def cleanup_estimate(self, host: str) -> CommandResult:
        script = r'''
$targets = @($env:TEMP,"$env:SystemRoot\Temp","$env:SystemRoot\SoftwareDistribution\Download")
$total = 0
$items = @()
foreach($path in $targets){
  if(Test-Path $path){
    $bytes = (Get-ChildItem $path -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
    if($bytes -eq $null){$bytes=0}
    $total += $bytes
    $items += [pscustomobject]@{Path=$path;SizeGB=[math]::Round($bytes/1GB,2)}
  }
}
[pscustomobject]@{RecoverableGB=[math]::Round($total/1GB,2);Areas=$items}
'''
        return self.executor.execute_powershell_json(host, script, timeout=180)

    def cleanup_safe(self, host: str) -> CommandResult:
        script = r'''
$targets = @($env:TEMP,"$env:SystemRoot\Temp")
foreach($path in $targets){
  if(Test-Path $path){ Get-ChildItem $path -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue }
}
try { Clear-RecycleBin -Force -ErrorAction SilentlyContinue } catch {}
'Limpeza segura concluída.'
'''
        return self.executor.execute_remote_powershell_with_fallback(host, script, timeout=180)
