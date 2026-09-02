from __future__ import annotations

from typing import Any, Dict

from core.executor import RemoteExecutor
from core.result import CommandResult


class DiagnosticsModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def preflight(self, host: str) -> Dict[str, Any]:
        ip = self.executor.resolve_host(host)
        ping = self.executor.ping(host)
        admin = self.executor.test_admin_share(host) if ping.success else None
        winrm = self.executor.test_winrm(host) if ping.success else None
        info = self.system_info(host) if ping.success else None

        return {
            "host": host,
            "ip": ip,
            "online": ping.success,
            "admin_share": bool(admin and admin.success),
            "winrm": bool(winrm and winrm.success),
            "system": info.data if info and info.success else None,
            "errors": [
                r.stderr
                for r in (ping, admin, winrm, info)
                if r is not None and not r.success and r.stderr
            ],
        }

    def system_info(self, host: str) -> CommandResult:
        script = r'''
$os = Get-CimInstance Win32_OperatingSystem
$cs = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$uptime = (Get-Date) - $os.LastBootUpTime
[pscustomobject]@{
    Hostname = $env:COMPUTERNAME
    User = (Get-CimInstance Win32_ComputerSystem).UserName
    Manufacturer = $cs.Manufacturer
    Model = $cs.Model
    Serial = $bios.SerialNumber
    BIOS = $bios.SMBIOSBIOSVersion
    OS = $os.Caption
    OSVersion = $os.Version
    Build = $os.BuildNumber
    CPU = $cpu.Name
    RAM_GB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
    FreeRAM_GB = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
    DiskC_GB = [math]::Round($disk.Size / 1GB, 2)
    DiskCFree_GB = [math]::Round($disk.FreeSpace / 1GB, 2)
    Uptime = ('{0}d {1}h {2}m' -f $uptime.Days,$uptime.Hours,$uptime.Minutes)
}
'''
        return self.executor.execute_powershell_json(host, script)

    def event_errors(self, host: str, hours: int = 24, limit: int = 20) -> CommandResult:
        script = f'''
$start = (Get-Date).AddHours(-{int(hours)})
Get-WinEvent -FilterHashtable @{{LogName='System'; StartTime=$start; Level=1,2}} -ErrorAction SilentlyContinue |
    Select-Object -First {int(limit)} TimeCreated, Id, ProviderName, LevelDisplayName, Message
'''
        return self.executor.execute_powershell_json(host, script)

    def pending_reboot(self, host: str) -> CommandResult:
        script = r'''
$paths = @(
 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
)
$pending = $false
foreach($p in $paths){ if(Test-Path $p){$pending=$true} }
[pscustomobject]@{ PendingReboot = $pending }
'''
        return self.executor.execute_powershell_json(host, script)
