from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class WorkstationToolsModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def certificates(self, host: str) -> CommandResult:
        script = r'''
$now=Get-Date
$certs=Get-ChildItem Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Select-Object Subject,Issuer,NotBefore,NotAfter,Thumbprint,HasPrivateKey,@{n='DaysToExpire';e={[math]::Floor(($_.NotAfter-$now).TotalDays)}}
$expired=@($certs|Where-Object {$_.DaysToExpire -lt 0})
$soon=@($certs|Where-Object {$_.DaysToExpire -ge 0 -and $_.DaysToExpire -le 30})
[pscustomobject]@{Certificates=$certs;Expired=$expired;ExpiringSoon=$soon}
'''
        return self.executor.execute_powershell_json(host, script, timeout=120)

    def mapped_drives(self, host: str) -> CommandResult:
        script = r'''
Get-CimInstance Win32_LogicalDisk -Filter "DriveType=4" | Select-Object DeviceID,ProviderName,VolumeName,Status
'''
        return self.executor.execute_powershell_json(host, script)

    def local_shares(self, host: str) -> CommandResult:
        script = r'''
Get-SmbShare -ErrorAction SilentlyContinue | Select-Object Name,Path,Description,Special,Temporary,FolderEnumerationMode
'''
        return self.executor.execute_powershell_json(host, script)

    def proxy(self, host: str) -> CommandResult:
        script = r'''
$winhttp = (netsh winhttp show proxy | Out-String).Trim()
$inet = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -ErrorAction SilentlyContinue
[pscustomobject]@{WinHTTP=$winhttp;ProxyEnable=$inet.ProxyEnable;ProxyServer=$inet.ProxyServer;AutoConfigURL=$inet.AutoConfigURL}
'''
        return self.executor.execute_powershell_json(host, script)

    def activation(self, host: str) -> CommandResult:
        script = r'''
$lic=Get-CimInstance SoftwareLicensingProduct -Filter "Name like 'Windows%' and PartialProductKey is not null" -ErrorAction SilentlyContinue | Select-Object -First 1 Name,Description,LicenseStatus,PartialProductKey,GracePeriodRemaining
$os=Get-CimInstance Win32_OperatingSystem
[pscustomobject]@{OS=$os.Caption;Build=$os.BuildNumber;License=$lic}
'''
        return self.executor.execute_powershell_json(host, script)

    def logons(self, host: str, hours: int = 24) -> CommandResult:
        hours=max(1,min(int(hours),168))
        script=f'''
$start=(Get-Date).AddHours(-{hours})
Get-WinEvent -FilterHashtable @{{LogName='Security';Id=4624;StartTime=$start}} -ErrorAction SilentlyContinue | Select-Object -First 100 TimeCreated,Id,Message
'''
        return self.executor.execute_powershell_json(host, script, timeout=180)
