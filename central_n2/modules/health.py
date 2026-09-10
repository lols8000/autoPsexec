from __future__ import annotations

from typing import Any

from core.evaluation import (
    EvaluationPolicy,
    EvaluationState,
    evaluate_snapshot,
    score_checks,
)
from core.executor import RemoteExecutor
from core.result import CommandResult


def calculate_health_score(
    data: dict[str, Any],
    *,
    min_free_disk_percent: int = 15,
    max_uptime_days: int = 30,
    baseline: dict[str, Any] | None = None,
) -> dict[str, Any]:
    values = dict(baseline or {})
    values.setdefault("min_disk_free_percent", min_free_disk_percent)
    values.setdefault("max_uptime_days", max_uptime_days)
    policy = EvaluationPolicy.from_mapping(values)
    checks = evaluate_snapshot(data, policy)
    summary = score_checks(checks)
    findings = [
        {
            "severity": check.severity,
            "message": check.message,
            "penalty": check.penalty,
            "key": check.key,
            "state": check.state.value,
        }
        for check in checks
        if check.state in {EvaluationState.FAIL, EvaluationState.UNKNOWN}
    ]
    return {
        "score": summary["health_score"],
        "overall_state": summary["overall_state"],
        "unknown": summary["unknown"],
        "findings": findings,
    }


class HealthModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def snapshot(self, host: str) -> CommandResult:
        script = r"""
$os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
$cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction SilentlyContinue
$cpu = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue |
    Measure-Object LoadPercentage -Average

$pending = $false
$keys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
)
foreach ($key in $keys) {
    if (Test-Path $key) { $pending = $true }
}
$pfro = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
if ($pfro) { $pending = $true }

$autoStopped = $null
try {
    $autoStopped = @(
        Get-CimInstance Win32_Service -ErrorAction Stop |
            Where-Object {
                $_.StartMode -eq 'Auto' -and $_.State -ne 'Running'
            }
    ).Count
} catch {}

$defender = $null
try {
    $defender = [bool](
        Get-MpComputerStatus -ErrorAction Stop
    ).AntivirusEnabled
} catch {}

$firewall = $null
try {
    $firewall = [bool](
        @(
            Get-NetFirewallProfile -ErrorAction Stop |
                Where-Object Enabled
        ).Count -gt 0
    )
} catch {}

$glpi = Get-Service -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -match '^glpi' -or $_.DisplayName -match 'GLPI'
    } |
    Select-Object -First 1
$glpiInstalled = [bool]$glpi
$glpiRunning = if ($glpi) { $glpi.Status -eq 'Running' } else { $false }

$bitlocker = $null
try {
    $bitlocker = (
        Get-BitLockerVolume -MountPoint 'C:' -ErrorAction Stop
    ).ProtectionStatus -eq 'On'
} catch {}

$tpmReady = $null
try {
    $tpm = Get-Tpm -ErrorAction Stop
    $tpmReady = [bool]($tpm.TpmPresent -and $tpm.TpmReady)
} catch {}

$secureBoot = $null
try {
    $secureBoot = Confirm-SecureBootUEFI -ErrorAction Stop
} catch {}

[pscustomobject]@{
    Hostname = $env:COMPUTERNAME
    User = $cs.UserName
    OS = $os.Caption
    Build = $os.BuildNumber
    Manufacturer = $cs.Manufacturer
    Model = $cs.Model
    CPUPercent = if ($cpu.Average -ne $null) {
        [math]::Round($cpu.Average, 1)
    } else { $null }
    RAMUsedPercent = if ($os.TotalVisibleMemorySize) {
        [math]::Round(
            (1 - ($os.FreePhysicalMemory / $os.TotalVisibleMemorySize)) * 100,
            1
        )
    } else { $null }
    DiskFreeGB = if ($disk) {
        [math]::Round($disk.FreeSpace / 1GB, 1)
    } else { $null }
    DiskFreePercent = if ($disk -and $disk.Size) {
        [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1)
    } else { $null }
    UptimeDays = if ($os.LastBootUpTime) {
        [math]::Floor(((Get-Date) - $os.LastBootUpTime).TotalDays)
    } else { $null }
    PendingReboot = $pending
    StoppedAutoServices = $autoStopped
    DefenderEnabled = $defender
    FirewallEnabled = $firewall
    GlpiInstalled = $glpiInstalled
    GlpiRunning = $glpiRunning
    BitLockerProtected = $bitlocker
    TPMReady = $tpmReady
    SecureBoot = $secureBoot
}
"""
        return self.executor.execute_powershell_json(
            host,
            script,
            timeout=90,
        )
