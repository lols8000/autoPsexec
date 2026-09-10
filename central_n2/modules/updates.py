from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class UpdatesModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def status(self, host: str) -> CommandResult:
        script = r"""
$history = Get-HotFix -ErrorAction SilentlyContinue |
    Sort-Object InstalledOn -Descending |
    Select-Object -First 20 HotFixID,Description,InstalledOn,InstalledBy

$wu = $null
try {
    $session = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $result = $searcher.Search("IsInstalled=0 and IsHidden=0")
    $pending = @(
        $result.Updates | ForEach-Object {
            [pscustomobject]@{
                Title = $_.Title
                KB = ($_.KBArticleIDs -join ',')
                RebootRequired = $_.RebootRequired
            }
        }
    )
    $wu = [pscustomobject]@{
        PendingCount = $result.Updates.Count
        Pending = $pending
        Error = $null
    }
} catch {
    $wu = [pscustomobject]@{
        PendingCount = $null
        Pending = @()
        Error = $_.Exception.Message
    }
}
[pscustomobject]@{History=$history;WindowsUpdate=$wu}
"""
        return self.executor.execute_powershell_json(host, script, timeout=180)

    def trigger_scan(self, host: str) -> CommandResult:
        script = r"""
$uso = Join-Path $env:SystemRoot 'System32\UsoClient.exe'
if (-not (Test-Path $uso)) { throw 'UsoClient.exe não encontrado.' }
Start-Process $uso -ArgumentList 'StartScan' -WindowStyle Hidden
'Busca iniciada.'
"""
        return self.executor.execute_mutating_powershell(
            host,
            script,
            timeout=60,
        )

    def reset_components(self, host: str) -> CommandResult:
        script = r"""
$serviceNames = @('bits','wuauserv','cryptsvc')
$original = @{}
$renamed = @()
$operationError = $null

foreach ($name in $serviceNames) {
    $svc = Get-Service -Name $name -ErrorAction Stop
    $original[$name] = $svc.Status.ToString()
}

try {
    foreach ($name in $serviceNames) {
        if ($original[$name] -eq 'Running') {
            Stop-Service -Name $name -Force -ErrorAction Stop
        }
    }

    $stamp = Get-Date -Format yyyyMMddHHmmss
    $softwareDistribution = "$env:SystemRoot\SoftwareDistribution"
    $catroot2 = "$env:SystemRoot\System32\catroot2"

    if (Test-Path $softwareDistribution) {
        $newName = "SoftwareDistribution.$stamp.bak"
        Rename-Item $softwareDistribution $newName -ErrorAction Stop
        $renamed += [pscustomobject]@{
            Original=$softwareDistribution
            Backup=(Join-Path $env:SystemRoot $newName)
        }
    }

    if (Test-Path $catroot2) {
        $newName = "catroot2.$stamp.bak"
        Rename-Item $catroot2 $newName -ErrorAction Stop
        $renamed += [pscustomobject]@{
            Original=$catroot2
            Backup=(Join-Path "$env:SystemRoot\System32" $newName)
        }
    }
} catch {
    $operationError = $_
} finally {
    foreach ($name in $serviceNames) {
        if ($original[$name] -eq 'Running') {
            try {
                Start-Service -Name $name -ErrorAction Stop
            } catch {
                if (-not $operationError) { $operationError = $_ }
            }
        }
    }
}

$states = @(
    Get-Service -Name $serviceNames -ErrorAction SilentlyContinue |
        Select-Object Name,Status,StartType
)

if ($operationError) {
    throw $operationError
}

[pscustomobject]@{
    Renamed = $renamed
    Services = $states
    RestoredOriginalRunningServices = [bool](
        @(
            $states |
                Where-Object {
                    $original[$_.Name] -eq 'Running' -and $_.Status -ne 'Running'
                }
        ).Count -eq 0
    )
}
"""
        result = self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=180,
        )
        result.metadata["transactional_cleanup"] = True
        return result
