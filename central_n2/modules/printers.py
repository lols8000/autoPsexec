from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import (
    quote_powershell_literal,
    validate_safe_name,
    validate_unc_path,
)


class PrintersModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def list(self, host: str) -> CommandResult:
        script = r"""
Get-Printer -ErrorAction SilentlyContinue |
    Select-Object Name,DriverName,PortName,PrinterStatus,WorkOffline,Shared,Published,Default
"""
        return self.executor.execute_powershell_json(host, script)

    def list_printers(self, host: str) -> CommandResult:
        return self.list(host)

    def printer_status(
        self,
        host: str,
        printer_name: str,
    ) -> CommandResult:
        safe = quote_powershell_literal(
            validate_safe_name(printer_name, label="Impressora")
        )
        script = f"""
$p = Get-Printer -Name {safe} -ErrorAction SilentlyContinue
[pscustomobject]@{{
    Exists = [bool]$p
    Name = if ($p) {{ $p.Name }} else {{ {safe} }}
    PrinterStatus = if ($p) {{ $p.PrinterStatus.ToString() }} else {{ $null }}
    DriverName = if ($p) {{ $p.DriverName }} else {{ $null }}
    PortName = if ($p) {{ $p.PortName }} else {{ $null }}
}}
"""
        return self.executor.execute_powershell_json(host, script)

    def queue(
        self,
        host: str,
        printer_name: str | None = None,
    ) -> CommandResult:
        if printer_name:
            safe = quote_powershell_literal(
                validate_safe_name(printer_name, label="Impressora")
            )
            script = (
                f"Get-PrintJob -PrinterName {safe} -ErrorAction SilentlyContinue | "
                "Select-Object PrinterName,ID,DocumentName,UserName,JobStatus,SubmittedTime,Size"
            )
        else:
            script = r"""
Get-Printer -ErrorAction SilentlyContinue | ForEach-Object {
    Get-PrintJob -PrinterName $_.Name -ErrorAction SilentlyContinue
} | Select-Object PrinterName,ID,DocumentName,UserName,JobStatus,SubmittedTime,Size
"""
        return self.executor.execute_powershell_json(host, script)

    def spooler_status(self, host: str) -> CommandResult:
        return self.executor.execute_powershell_json(
            host,
            (
                "Get-Service Spooler | "
                "Select-Object Name,@{n='Status';e={$_.Status.ToString()}},StartType"
            ),
        )

    def restart_spooler(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_powershell_json(
            host,
            (
                "Restart-Service Spooler -Force -ErrorAction Stop; "
                "Get-Service Spooler | "
                "Select-Object Name,@{n='Status';e={$_.Status.ToString()}},StartType"
            ),
        )

    def clear_queue(
        self,
        host: str,
        printer_name: str | None = None,
    ) -> CommandResult:
        if printer_name:
            safe = quote_powershell_literal(
                validate_safe_name(printer_name, label="Impressora")
            )
            script = f"""
$jobs = @(Get-PrintJob -PrinterName {safe} -ErrorAction SilentlyContinue)
$count = $jobs.Count
$jobs | Remove-PrintJob -ErrorAction Stop
[pscustomobject]@{{Printer={safe};RemovedJobs=$count}}
"""
        else:
            script = r"""
$jobs = @(
    Get-Printer -ErrorAction SilentlyContinue | ForEach-Object {
        Get-PrintJob -PrinterName $_.Name -ErrorAction SilentlyContinue
    }
)
$count = $jobs.Count
$jobs | Remove-PrintJob -ErrorAction Stop
[pscustomobject]@{Printer='*';RemovedJobs=$count}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=90,
        )

    def add_connection(
        self,
        host: str,
        connection_name: str,
    ) -> CommandResult:
        safe = quote_powershell_literal(validate_unc_path(connection_name))
        script = f"""
Add-Printer -ConnectionName {safe} -ErrorAction Stop
$p = Get-Printer -Name {safe} -ErrorAction SilentlyContinue
[pscustomobject]@{{
    Added = [bool]$p
    Name = if ($p) {{ $p.Name }} else {{ {safe} }}
}}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=180,
        )

    def remove_printer(
        self,
        host: str,
        printer_name: str,
    ) -> CommandResult:
        safe = quote_powershell_literal(
            validate_safe_name(printer_name, label="Impressora")
        )
        script = f"""
Remove-Printer -Name {safe} -ErrorAction Stop
$remaining = Get-Printer -Name {safe} -ErrorAction SilentlyContinue
[pscustomobject]@{{Removed=[bool](-not $remaining);Name={safe}}}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=120,
        )
