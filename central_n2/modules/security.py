from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class SecurityModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def status(self, host: str) -> CommandResult:
        script = r'''
$def = $null
try {
  $m = Get-MpComputerStatus -ErrorAction Stop
  $def = [pscustomobject]@{
    AntivirusEnabled=$m.AntivirusEnabled
    RealTimeProtectionEnabled=$m.RealTimeProtectionEnabled
    AntivirusSignatureLastUpdated=$m.AntivirusSignatureLastUpdated
    QuickScanAge=$m.QuickScanAge
    FullScanAge=$m.FullScanAge
  }
} catch {}
$fw = Get-NetFirewallProfile -ErrorAction SilentlyContinue | Select-Object Name,Enabled,DefaultInboundAction,DefaultOutboundAction
$bitlocker = @()
try { $bitlocker = Get-BitLockerVolume -ErrorAction Stop | Select-Object MountPoint,VolumeStatus,ProtectionStatus,EncryptionPercentage } catch {}
$tpm = $null
try { $t = Get-Tpm -ErrorAction Stop; $tpm = [pscustomobject]@{Present=$t.TpmPresent;Ready=$t.TpmReady;Enabled=$t.TpmEnabled;Activated=$t.TpmActivated} } catch {}
$secureBoot = $null
try { $secureBoot = Confirm-SecureBootUEFI -ErrorAction Stop } catch {}
$rdp = $null
try { $rdp = ((Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server').fDenyTSConnections -eq 0) } catch {}
$smb1 = $null
try { $smb1 = (Get-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -ErrorAction Stop).State -eq 'Enabled' } catch {}
[pscustomobject]@{
 Defender=$def
 Firewall=$fw
 BitLocker=$bitlocker
 TPM=$tpm
 SecureBoot=$secureBoot
 RDPEnabled=$rdp
 SMB1Enabled=$smb1
 UAC=(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction SilentlyContinue).EnableLUA -eq 1
}
'''
        return self.executor.execute_powershell_json(host, script, timeout=120)

    def threats(self, host: str) -> CommandResult:
        script = r'''
try {
  Get-MpThreatDetection -ErrorAction Stop |
    Sort-Object InitialDetectionTime -Descending |
    Select-Object -First 25 ThreatName,InitialDetectionTime,LastThreatStatusChangeTime,ActionSuccess,CurrentThreatExecutionStatusID
} catch { @() }
'''
        return self.executor.execute_powershell_json(host, script)
