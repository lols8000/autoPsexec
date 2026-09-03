from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class PerformanceModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def snapshot(self, host: str, samples: int = 8, interval: int = 1) -> CommandResult:
        samples = max(3, min(int(samples), 30))
        interval = max(1, min(int(interval), 5))
        script = f'''
$cpu=@(); $disk=@(); $net=@()
1..{samples} | ForEach-Object {{
  $cpu += (Get-CimInstance Win32_Processor | Measure-Object LoadPercentage -Average).Average
  $disk += (Get-Counter '\\PhysicalDisk(_Total)\\% Disk Time' -ErrorAction SilentlyContinue).CounterSamples.CookedValue
  $net += ((Get-Counter '\\Network Interface(*)\\Bytes Total/sec' -ErrorAction SilentlyContinue).CounterSamples | Measure-Object CookedValue -Sum).Sum
  Start-Sleep -Seconds {interval}
}}
$os=Get-CimInstance Win32_OperatingSystem
$topCpu=Get-Process -ErrorAction SilentlyContinue | Sort-Object CPU -Descending | Select-Object -First 8 Name,Id,CPU,@{{n='RAMMB';e={{[math]::Round($_.WorkingSet64/1MB,1)}}}}
$topMem=Get-Process -ErrorAction SilentlyContinue | Sort-Object WorkingSet64 -Descending | Select-Object -First 8 Name,Id,@{{n='RAMMB';e={{[math]::Round($_.WorkingSet64/1MB,1)}}}},CPU
[pscustomobject]@{{
 CPUAverage=[math]::Round(($cpu|Measure-Object -Average).Average,1)
 CPUMax=[math]::Round(($cpu|Measure-Object -Maximum).Maximum,1)
 RAMUsedPercent=[math]::Round((1-($os.FreePhysicalMemory/$os.TotalVisibleMemorySize))*100,1)
 DiskAverage=[math]::Round(($disk|Measure-Object -Average).Average,1)
 NetworkMbps=[math]::Round((($net|Measure-Object -Average).Average*8/1MB),2)
 TopCPU=$topCpu
 TopMemory=$topMem
}}
'''
        return self.executor.execute_powershell_json(host, script, timeout=samples * interval + 90)
