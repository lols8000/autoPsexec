from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class PerformanceModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def snapshot(
        self,
        host: str,
        samples: int = 8,
        interval: int = 1,
    ) -> CommandResult:
        samples = max(3, min(int(samples), 30))
        interval = max(1, min(int(interval), 5))

        script = f"""
$cpu = @()
$disk = @()
$net = @()

1..{samples} | ForEach-Object {{
    $cpuValue = (
        Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue |
        Measure-Object LoadPercentage -Average
    ).Average
    if ($cpuValue -ne $null) {{ $cpu += $cpuValue }}

    $diskCounter = (
        Get-Counter '\\PhysicalDisk(_Total)\\% Disk Time' -ErrorAction SilentlyContinue
    ).CounterSamples.CookedValue
    if ($diskCounter -ne $null) {{ $disk += $diskCounter }}

    $netCounter = (
        (
            Get-Counter '\\Network Interface(*)\\Bytes Total/sec' -ErrorAction SilentlyContinue
        ).CounterSamples |
        Measure-Object CookedValue -Sum
    ).Sum
    if ($netCounter -ne $null) {{ $net += $netCounter }}

    Start-Sleep -Seconds {interval}
}}

$os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
$logical = [Math]::Max(1, [Environment]::ProcessorCount)

$before = @{{}}
Get-Process -ErrorAction SilentlyContinue | ForEach-Object {{
    if ($_.CPU -ne $null) {{
        $before[$_.Id] = [double]$_.CPU
    }}
}}

$cpuWindowSeconds = 1.0
Start-Sleep -Seconds $cpuWindowSeconds

$topCpuCurrent = @(
    Get-Process -ErrorAction SilentlyContinue | ForEach-Object {{
        if ($_.CPU -ne $null -and $before.ContainsKey($_.Id)) {{
            $delta = [double]$_.CPU - [double]$before[$_.Id]
            $percent = ($delta / $cpuWindowSeconds) * 100 / $logical
            if ($percent -lt 0) {{ $percent = 0 }}
            [pscustomobject]@{{
                Name = $_.Name
                Id = $_.Id
                CPUCurrentPercent = [math]::Round($percent, 1)
                RAMMB = [math]::Round($_.WorkingSet64 / 1MB, 1)
            }}
        }}
    }} |
    Sort-Object CPUCurrentPercent -Descending |
    Select-Object -First 10
)

$topMem = @(
    Get-Process -ErrorAction SilentlyContinue |
    Sort-Object WorkingSet64 -Descending |
    Select-Object -First 10 Name,Id,@{{
        n='RAMMB';e={{[math]::Round($_.WorkingSet64 / 1MB, 1)}}
    }}
)

[pscustomobject]@{{
    CPUAverage = if ($cpu.Count) {{
        [math]::Round(($cpu | Measure-Object -Average).Average, 1)
    }} else {{ $null }}
    CPUMax = if ($cpu.Count) {{
        [math]::Round(($cpu | Measure-Object -Maximum).Maximum, 1)
    }} else {{ $null }}
    RAMUsedPercent = if ($os.TotalVisibleMemorySize) {{
        [math]::Round(
            (1 - ($os.FreePhysicalMemory / $os.TotalVisibleMemorySize)) * 100,
            1
        )
    }} else {{ $null }}
    DiskAverage = if ($disk.Count) {{
        [math]::Round(($disk | Measure-Object -Average).Average, 1)
    }} else {{ $null }}
    NetworkMbps = if ($net.Count) {{
        [math]::Round(
            (($net | Measure-Object -Average).Average * 8 / 1MB),
            2
        )
    }} else {{ $null }}
    TopCPU = $topCpuCurrent
    TopMemory = $topMem
}}
"""
        return self.executor.execute_powershell_json(
            host,
            script,
            timeout=samples * interval + 100,
        )
