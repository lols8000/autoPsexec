from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import (
    quote_cmd_argument,
    quote_powershell_literal,
    validate_inf_name,
    validate_safe_name,
    validate_windows_path,
)


class DevicesModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def problem_devices(self, host: str) -> CommandResult:
        script = r"""
Get-CimInstance Win32_PnPEntity |
    Where-Object { $_.ConfigManagerErrorCode -ne 0 } |
    Select-Object Name,PNPDeviceID,Manufacturer,Status,ConfigManagerErrorCode |
    Sort-Object ConfigManagerErrorCode,Name
"""
        return self.executor.execute_powershell_json(host, script, timeout=120)

    def device_status(
        self,
        host: str,
        instance_id: str,
    ) -> CommandResult:
        safe = quote_powershell_literal(
            validate_safe_name(instance_id, label="InstanceId")
        )
        script = f"""
$d = Get-PnpDevice -InstanceId {safe} -ErrorAction SilentlyContinue
[pscustomobject]@{{
    Exists = [bool]$d
    FriendlyName = if ($d) {{ $d.FriendlyName }} else {{ $null }}
    InstanceId = {safe}
    Status = if ($d) {{ $d.Status.ToString() }} else {{ $null }}
    Problem = if ($d) {{ $d.Problem }} else {{ $null }}
}}
"""
        return self.executor.execute_powershell_json(host, script)

    def drivers(self, host: str) -> CommandResult:
        script = r"""
Get-CimInstance Win32_PnPSignedDriver |
    Where-Object { $_.DeviceName -or $_.InfName -or $_.Manufacturer } |
    ForEach-Object {
        [pscustomobject]@{
            DeviceName = $_.DeviceName
            Manufacturer = $_.Manufacturer
            DriverVersion = $_.DriverVersion
            DriverDate = if ($_.DriverDate) {
                ([datetime]$_.DriverDate).ToString('yyyy-MM-dd')
            } else { $null }
            IsSigned = $_.IsSigned
            InfName = $_.InfName
        }
    } |
    Group-Object DeviceName,Manufacturer,DriverVersion,DriverDate,IsSigned,InfName |
    ForEach-Object {
        $item = $_.Group[0]
        [pscustomobject]@{
            DeviceName = $item.DeviceName
            Manufacturer = $item.Manufacturer
            DriverVersion = $item.DriverVersion
            DriverDate = $item.DriverDate
            IsSigned = $item.IsSigned
            InfName = $item.InfName
            Count = $_.Count
        }
    } |
    Sort-Object DeviceName,Manufacturer
"""
        result = self.executor.execute_powershell_json(host, script, timeout=180)
        result.metadata["view"] = "drivers"
        return result

    def usb_devices(self, host: str) -> CommandResult:
        script = r"""
Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue |
    Where-Object { $_.InstanceId -like 'USB*' } |
    Select-Object Class,FriendlyName,InstanceId,Status,Problem
"""
        return self.executor.execute_powershell_json(host, script, timeout=120)

    def set_device_state(
        self,
        host: str,
        instance_id: str,
        *,
        enabled: bool,
    ) -> CommandResult:
        safe = quote_powershell_literal(
            validate_safe_name(instance_id, label="InstanceId")
        )
        cmd = "Enable-PnpDevice" if enabled else "Disable-PnpDevice"
        script = f"""
{cmd} -InstanceId {safe} -Confirm:$false -ErrorAction Stop
Start-Sleep -Milliseconds 500
$d = Get-PnpDevice -InstanceId {safe} -ErrorAction Stop
[pscustomobject]@{{
    InstanceId = $d.InstanceId
    FriendlyName = $d.FriendlyName
    Status = $d.Status.ToString()
    Problem = $d.Problem
}}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=120,
        )

    def rescan(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(
            host,
            "pnputil /scan-devices",
            timeout=180,
        )

    def install_driver_inf(
        self,
        host: str,
        path: str,
    ) -> CommandResult:
        safe_path = quote_cmd_argument(validate_windows_path(path))
        return self.executor.execute_mutating_cmd(
            host,
            f"pnputil /add-driver {safe_path} /install",
            timeout=600,
        )

    def remove_driver_package(
        self,
        host: str,
        inf_name: str,
    ) -> CommandResult:
        safe_inf = quote_cmd_argument(validate_inf_name(inf_name))
        return self.executor.execute_mutating_cmd(
            host,
            f"pnputil /delete-driver {safe_inf} /uninstall /force",
            timeout=600,
        )

    def export_drivers(
        self,
        host: str,
        path: str = r"C:\CentralN2\Drivers",
    ) -> CommandResult:
        safe_path = quote_cmd_argument(validate_windows_path(path))
        return self.executor.execute_mutating_cmd(
            host,
            f"pnputil /export-driver * {safe_path}",
            timeout=1800,
        )
