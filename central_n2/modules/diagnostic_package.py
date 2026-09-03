from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from core.executor import RemoteExecutor


class DiagnosticPackageModule:
    def __init__(self, executor: RemoteExecutor, reports_dir: Path) -> None:
        self.executor = executor
        self.reports_dir = reports_dir

    def collect(self, host: str) -> Path:
        script = r'''
$os = Get-CimInstance Win32_OperatingSystem
$cs = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS
$disk = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | Select-Object DeviceID,Size,FreeSpace
$net = Get-NetIPConfiguration -ErrorAction SilentlyContinue | Select-Object InterfaceAlias,IPv4Address,IPv4DefaultGateway,DNSServer
$services = Get-CimInstance Win32_Service | Where-Object {$_.StartMode -eq 'Auto' -and $_.State -ne 'Running'} | Select-Object Name,DisplayName,State,StartMode
$events = Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=(Get-Date).AddHours(-24)} -ErrorAction SilentlyContinue | Select-Object -First 50 TimeCreated,Id,ProviderName,LevelDisplayName,Message
$printers = Get-Printer -ErrorAction SilentlyContinue | Select-Object Name,DriverName,PortName,PrinterStatus
$processes = Get-Process | Sort-Object WorkingSet -Descending | Select-Object -First 25 Name,Id,CPU,WorkingSet
[pscustomobject]@{
 Computer=[pscustomobject]@{Hostname=$env:COMPUTERNAME;Manufacturer=$cs.Manufacturer;Model=$cs.Model;User=$cs.UserName;Serial=$bios.SerialNumber}
 OS=[pscustomobject]@{Caption=$os.Caption;Build=$os.BuildNumber;LastBootUpTime=$os.LastBootUpTime}
 Disks=$disk
 Network=$net
 StoppedAutomaticServices=$services
 RecentCriticalEvents=$events
 Printers=$printers
 TopProcesses=$processes
}
'''
        result = self.executor.execute_powershell_json(host, script, timeout=180)
        if not result.success:
            raise RuntimeError(result.stderr or "Falha ao coletar diagnóstico")
        self.reports_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = self.reports_dir / f"{host}_{stamp}.json"
        target.write_text(json.dumps(result.data, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        return target
