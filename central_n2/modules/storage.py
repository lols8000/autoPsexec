from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class StorageModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def physical_disks(self, host: str) -> CommandResult:
        script = r'''
$items=@()
try {
  $items=Get-PhysicalDisk -ErrorAction Stop | Select-Object FriendlyName,SerialNumber,MediaType,BusType,HealthStatus,OperationalStatus,@{n='SizeGB';e={[math]::Round($_.Size/1GB,1)}}
} catch {
  $items=Get-CimInstance Win32_DiskDrive | Select-Object Model,SerialNumber,InterfaceType,Status,@{n='SizeGB';e={[math]::Round($_.Size/1GB,1)}}
}
$items
'''
        return self.executor.execute_powershell_json(host, script, timeout=120)

    def battery(self, host: str) -> CommandResult:
        script = r'''
$bat=Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue | Select-Object Name,Status,EstimatedChargeRemaining,BatteryStatus,EstimatedRunTime
$designed=$null;$full=$null;$cycles=$null
try {
  $ns='root\wmi'
  $static=Get-CimInstance -Namespace $ns -ClassName BatteryStaticData -ErrorAction Stop | Select-Object -First 1
  $fullData=Get-CimInstance -Namespace $ns -ClassName BatteryFullChargedCapacity -ErrorAction Stop | Select-Object -First 1
  $cycle=Get-CimInstance -Namespace $ns -ClassName BatteryCycleCount -ErrorAction SilentlyContinue | Select-Object -First 1
  $designed=$static.DesignedCapacity;$full=$fullData.FullChargedCapacity;$cycles=$cycle.CycleCount
} catch {}
[pscustomobject]@{Battery=$bat;DesignedCapacity=$designed;FullChargeCapacity=$full;HealthPercent=if($designed -and $full){[math]::Round(($full/$designed)*100,1)}else{$null};CycleCount=$cycles}
'''
        return self.executor.execute_powershell_json(host, script, timeout=120)
