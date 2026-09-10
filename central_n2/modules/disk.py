from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class DiskModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def usage(self, host: str) -> CommandResult:
        script = r"""
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
if (-not $disk -or -not $disk.Size) {
    throw 'Não foi possível consultar o volume C:.'
}
[pscustomobject]@{
    DeviceID = 'C:'
    SizeGB = [math]::Round($disk.Size / 1GB, 1)
    FreeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
    FreePercent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1)
}
"""
        return self.executor.execute_powershell_json(host, script)

    def space(self, host: str) -> CommandResult:
        return self.usage(host)

    def top_user_profiles(self, host: str) -> CommandResult:
        script = r"""
Get-CimInstance Win32_UserProfile |
    Where-Object { -not $_.Special -and $_.LocalPath } |
    ForEach-Object {
        $path = $_.LocalPath
        $size = $null
        try {
            $bytes = (
                Get-ChildItem -LiteralPath $path -File -Recurse -Force -ErrorAction SilentlyContinue |
                Measure-Object Length -Sum
            ).Sum
            if ($bytes -ne $null) {
                $size = [math]::Round($bytes / 1GB, 2)
            }
        } catch {}
        [pscustomobject]@{
            Profile = Split-Path $path -Leaf
            Path = $path
            Loaded = $_.Loaded
            LastUseTime = $_.LastUseTime
            SizeGB = $size
        }
    } |
    Sort-Object SizeGB -Descending
"""
        result = self.executor.execute_powershell_json(host, script, timeout=600)
        result.metadata["heavy_read"] = True
        return result

    def profile_sizes(self, host: str) -> CommandResult:
        return self.top_user_profiles(host)

    def cleanup_estimate(self, host: str) -> CommandResult:
        script = r"""
$targets = @(
    [pscustomobject]@{Name='UserTemp';Path=$env:TEMP},
    [pscustomobject]@{Name='WindowsTemp';Path="$env:SystemRoot\Temp"}
)
$total = 0
$items = @()
foreach ($target in $targets) {
    $bytes = 0
    if (Test-Path $target.Path) {
        $value = (
            Get-ChildItem $target.Path -Recurse -Force -File -ErrorAction SilentlyContinue |
            Measure-Object Length -Sum
        ).Sum
        if ($value -ne $null) { $bytes = $value }
    }
    $total += $bytes
    $items += [pscustomobject]@{
        Name = $target.Name
        Path = $target.Path
        SizeGB = [math]::Round($bytes / 1GB, 2)
    }
}
[pscustomobject]@{
    RecoverableGB = [math]::Round($total / 1GB, 2)
    Areas = $items
    IncludesRecycleBin = $false
    IncludesWindowsUpdateCache = $false
}
"""
        return self.executor.execute_powershell_json(host, script, timeout=180)

    def cleanup_safe(self, host: str) -> CommandResult:
        script = r"""
$targets = @($env:TEMP, "$env:SystemRoot\Temp")
$before = 0
$after = 0
foreach ($path in $targets) {
    if (Test-Path $path) {
        $bytes = (
            Get-ChildItem $path -Recurse -Force -File -ErrorAction SilentlyContinue |
            Measure-Object Length -Sum
        ).Sum
        if ($bytes) { $before += $bytes }
        Get-ChildItem $path -Force -ErrorAction SilentlyContinue |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        $bytesAfter = (
            Get-ChildItem $path -Recurse -Force -File -ErrorAction SilentlyContinue |
            Measure-Object Length -Sum
        ).Sum
        if ($bytesAfter) { $after += $bytesAfter }
    }
}
[pscustomobject]@{
    BeforeGB = [math]::Round($before / 1GB, 2)
    AfterGB = [math]::Round($after / 1GB, 2)
    RecoveredGB = [math]::Round(($before - $after) / 1GB, 2)
    RecycleBinTouched = $false
}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=180,
        )
