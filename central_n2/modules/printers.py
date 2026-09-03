from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class PrintersModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def list(self, host: str) -> CommandResult:
        script = r'''
Get-Printer -ErrorAction SilentlyContinue |
 Select-Object Name,DriverName,PortName,PrinterStatus,WorkOffline,Shared,Published,Default
'''
        return self.executor.execute_powershell_json(host, script)

    def list_printers(self, host: str) -> CommandResult:
        return self.list(host)

    def queue(self, host: str, printer_name: str | None = None) -> CommandResult:
        if printer_name:
            safe = printer_name.replace("'", "''")
            script = f"Get-PrintJob -PrinterName '{safe}' -ErrorAction SilentlyContinue | Select-Object PrinterName,ID,DocumentName,UserName,JobStatus,SubmittedTime,Size"
        else:
            script = r'''
Get-Printer -ErrorAction SilentlyContinue | ForEach-Object {
  Get-PrintJob -PrinterName $_.Name -ErrorAction SilentlyContinue
} | Select-Object PrinterName,ID,DocumentName,UserName,JobStatus,SubmittedTime,Size
'''
        return self.executor.execute_powershell_json(host, script)

    def restart_spooler(self, host: str) -> CommandResult:
        return self.executor.execute_remote_powershell_with_fallback(host, "Restart-Service Spooler -Force -ErrorAction Stop")

    def clear_queue(self, host: str, printer_name: str | None = None) -> CommandResult:
        if printer_name:
            safe = printer_name.replace("'", "''")
            script = f"Get-PrintJob -PrinterName '{safe}' -ErrorAction SilentlyContinue | Remove-PrintJob -ErrorAction Stop"
        else:
            script = r'''
Get-Printer -ErrorAction SilentlyContinue | ForEach-Object {
  Get-PrintJob -PrinterName $_.Name -ErrorAction SilentlyContinue | Remove-PrintJob -ErrorAction SilentlyContinue
}
'''
        return self.executor.execute_remote_powershell_with_fallback(host, script, timeout=90)
