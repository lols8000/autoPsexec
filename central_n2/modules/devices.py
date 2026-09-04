from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class DevicesModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def problem_devices(self, host: str) -> CommandResult:
        script = r'''
Get-CimInstance Win32_PnPEntity | Where-Object {$_.ConfigManagerErrorCode -ne 0} |
 Select-Object Name,PNPDeviceID,Manufacturer,Status,ConfigManagerErrorCode |
 Sort-Object ConfigManagerErrorCode,Name
'''
        return self.executor.execute_powershell_json(host, script, timeout=120)

    def drivers(self, host: str) -> CommandResult:
        script = r'''
Get-CimInstance Win32_PnPSignedDriver |
 Where-Object {$_.DeviceName -or $_.InfName -or $_.Manufacturer} |
 ForEach-Object {
    [pscustomobject]@{
        DeviceName    = $_.DeviceName
        Manufacturer  = $_.Manufacturer
        DriverVersion = $_.DriverVersion
        DriverDate    = if ($_.DriverDate) { ([datetime]$_.DriverDate).ToString('yyyy-MM-dd') } else { $null }
        IsSigned      = $_.IsSigned
        InfName       = $_.InfName
    }
 } |
 Group-Object DeviceName,Manufacturer,DriverVersion,DriverDate,IsSigned,InfName |
 ForEach-Object {
    $item = $_.Group[0]
    [pscustomobject]@{
        DeviceName    = $item.DeviceName
        Manufacturer  = $item.Manufacturer
        DriverVersion = $item.DriverVersion
        DriverDate    = $item.DriverDate
        IsSigned      = $item.IsSigned
        InfName       = $item.InfName
        Count         = $_.Count
    }
 } |
 Sort-Object DeviceName,Manufacturer
'''
        result = self.executor.execute_powershell_json(host, script, timeout=180)
        result.metadata["view"] = "drivers"
        return result

    def usb_devices(self, host: str) -> CommandResult:
        script = r'''
Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue |
 Where-Object {$_.InstanceId -like 'USB*'} |
 Select-Object Class,FriendlyName,InstanceId,Status,Problem
'''
        return self.executor.execute_powershell_json(host, script, timeout=120)

    def rescan(self, host: str) -> CommandResult:
        return self.executor.execute_cmd(host, "pnputil /scan-devices", timeout=180)

    def export_drivers(self, host: str, path: str = r"C:\CentralN2\Drivers") -> CommandResult:
        safe = path.replace('"', '')
        return self.executor.execute_cmd(host, f'pnputil /export-driver * "{safe}"', timeout=1800)
