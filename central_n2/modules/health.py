from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


def calculate_health_score(data: dict, *, min_free_disk_percent: int = 15, max_uptime_days: int = 30) -> dict:
    score = 100
    findings: list[dict] = []

    def add(severity: str, message: str, penalty: int) -> None:
        nonlocal score
        score -= penalty
        findings.append({"severity": severity, "message": message, "penalty": penalty})

    free = float(data.get("DiskFreePercent") or 0)
    uptime = int(data.get("UptimeDays") or 0)
    if free < 5:
        add("critical", f"Disco C: com apenas {free:.1f}% livre", 25)
    elif free < min_free_disk_percent:
        add("high", f"Disco C: abaixo do mínimo ({free:.1f}% livre)", 15)
    if uptime > max_uptime_days:
        add("medium", f"Uptime elevado: {uptime} dias", 8)
    if data.get("PendingReboot"):
        add("medium", "Reinicialização pendente", 8)
    if int(data.get("StoppedAutoServices") or 0) > 0:
        add("high", f"{data['StoppedAutoServices']} serviço(s) automático(s) parado(s)", 12)
    if data.get("DefenderEnabled") is False:
        add("critical", "Microsoft Defender desativado", 25)
    if data.get("FirewallEnabled") is False:
        add("high", "Firewall sem perfil ativo", 15)
    if data.get("GlpiRunning") is False:
        add("medium", "GLPI Agent não está em execução", 8)

    return {"score": max(0, score), "findings": findings}


class HealthModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def snapshot(self, host: str) -> CommandResult:
        script = r'''
$os = Get-CimInstance Win32_OperatingSystem
$cs = Get-CimInstance Win32_ComputerSystem
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$cpu = Get-CimInstance Win32_Processor | Measure-Object LoadPercentage -Average
$pending = $false
$keys = @(
 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
)
foreach($k in $keys){ if(Test-Path $k){ $pending = $true } }
$pfro = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
if($pfro){ $pending = $true }
$autoStopped = @(Get-CimInstance Win32_Service | Where-Object {$_.StartMode -eq 'Auto' -and $_.State -ne 'Running'}).Count
$defender = $null
try { $defender = [bool](Get-MpComputerStatus -ErrorAction Stop).AntivirusEnabled } catch {}
$fw = $null
try { $fw = [bool](@(Get-NetFirewallProfile -ErrorAction Stop | Where-Object Enabled).Count -gt 0) } catch {}
$glpi = Get-Service -Name 'glpi-agent','GLPI-Agent' -ErrorAction SilentlyContinue | Select-Object -First 1
[pscustomobject]@{
 Hostname=$env:COMPUTERNAME
 User=(Get-CimInstance Win32_ComputerSystem).UserName
 OS=$os.Caption
 Build=$os.BuildNumber
 Manufacturer=$cs.Manufacturer
 Model=$cs.Model
 CPUPercent=[math]::Round(($cpu.Average),1)
 RAMUsedPercent=[math]::Round((1-($os.FreePhysicalMemory/$os.TotalVisibleMemorySize))*100,1)
 DiskFreeGB=if($disk){[math]::Round($disk.FreeSpace/1GB,1)}else{$null}
 DiskFreePercent=if($disk){[math]::Round(($disk.FreeSpace/$disk.Size)*100,1)}else{$null}
 UptimeDays=[math]::Floor(((Get-Date)-$os.LastBootUpTime).TotalDays)
 PendingReboot=$pending
 StoppedAutoServices=$autoStopped
 DefenderEnabled=$defender
 FirewallEnabled=$fw
 GlpiRunning=if($glpi){$glpi.Status -eq 'Running'}else{$null}
}
'''
        return self.executor.execute_powershell_json(host, script, timeout=90)
