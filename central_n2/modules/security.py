from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class SecurityModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def status(self, host: str) -> CommandResult:
        script = r"""
$def = $null
try {
    $m = Get-MpComputerStatus -ErrorAction Stop
    $def = [pscustomobject]@{
        AntivirusEnabled = $m.AntivirusEnabled
        RealTimeProtectionEnabled = $m.RealTimeProtectionEnabled
        AntivirusSignatureLastUpdated = $m.AntivirusSignatureLastUpdated
        QuickScanAge = $m.QuickScanAge
        FullScanAge = $m.FullScanAge
    }
} catch {}

$fw = Get-NetFirewallProfile -ErrorAction SilentlyContinue |
    Select-Object Name,Enabled,DefaultInboundAction,DefaultOutboundAction

$bitlocker = @()
try {
    $bitlocker = Get-BitLockerVolume -ErrorAction Stop |
        Select-Object MountPoint,VolumeStatus,ProtectionStatus,EncryptionPercentage
} catch {}

$tpm = $null
try {
    $t = Get-Tpm -ErrorAction Stop
    $tpm = [pscustomobject]@{
        Present = $t.TpmPresent
        Ready = $t.TpmReady
        Enabled = $t.TpmEnabled
        Activated = $t.TpmActivated
    }
} catch {}

$secureBoot = $null
try { $secureBoot = Confirm-SecureBootUEFI -ErrorAction Stop } catch {}

$rdp = $null
try {
    $rdp = ((Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server').fDenyTSConnections -eq 0)
} catch {}

$smb1 = $null
try {
    $smb1 = (Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -ErrorAction Stop).State -eq 'Enabled'
} catch {}

$uac = $null
try {
    $uac = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction Stop).EnableLUA -eq 1
} catch {}

[pscustomobject]@{
    Defender = $def
    Firewall = $fw
    BitLocker = $bitlocker
    TPM = $tpm
    SecureBoot = $secureBoot
    RDPEnabled = $rdp
    SMB1Enabled = $smb1
    UAC = $uac
}
"""
        return self.executor.execute_powershell_json(host, script, timeout=120)

    def posture(self, host: str) -> CommandResult:
        return self.status(host)

    def threats(self, host: str) -> CommandResult:
        script = r"""
try {
    Get-MpThreatDetection -ErrorAction Stop |
        Sort-Object InitialDetectionTime -Descending |
        Select-Object -First 25 ThreatName,InitialDetectionTime,LastThreatStatusChangeTime,ActionSuccess,CurrentThreatExecutionStatusID
} catch { @() }
"""
        return self.executor.execute_powershell_json(host, script)

    def defender_status(self, host: str) -> CommandResult:
        script = r"""
$m = Get-MpComputerStatus -ErrorAction Stop
[pscustomobject]@{
    AntivirusEnabled = [bool]$m.AntivirusEnabled
    RealTimeProtectionEnabled = [bool]$m.RealTimeProtectionEnabled
    SignatureLastUpdated = $m.AntivirusSignatureLastUpdated
    QuickScanAge = $m.QuickScanAge
    FullScanAge = $m.FullScanAge
}
"""
        return self.executor.execute_powershell_json(host, script)

    def update_defender_signatures(self, host: str) -> CommandResult:
        script = r"""
Update-MpSignature -ErrorAction Stop
$m = Get-MpComputerStatus -ErrorAction Stop
[pscustomobject]@{
    Updated = $true
    SignatureLastUpdated = $m.AntivirusSignatureLastUpdated
}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=600,
        )

    def defender_scan(
        self,
        host: str,
        scan_type: str,
    ) -> CommandResult:
        allowed = {
            "QuickScan": "QuickScan",
            "FullScan": "FullScan",
        }
        value = allowed.get(scan_type)
        if not value:
            return CommandResult.failure(
                host,
                "defender_scan",
                "Tipo de varredura inválido.",
            )
        script = f"""
Start-MpScan -ScanType {value} -ErrorAction Stop
$m = Get-MpComputerStatus -ErrorAction Stop
[pscustomobject]@{{
    ScanType = '{value}'
    QuickScanAge = $m.QuickScanAge
    FullScanAge = $m.FullScanAge
    AntivirusEnabled = [bool]$m.AntivirusEnabled
}}
"""
        timeout = 1800 if value == "QuickScan" else 7200
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=timeout,
        )
