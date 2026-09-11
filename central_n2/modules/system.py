from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import (
    quote_powershell_literal,
    validate_process_name,
    validate_safe_name,
)


class SystemModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def sessions(self, host: str) -> CommandResult:
        return self.executor.execute_cmd(host, "quser")

    def send_message(self, host: str, message: str) -> CommandResult:
        safe = message.replace('"', "'").replace("\r", " ").replace("\n", " ")
        return self.executor.execute_mutating_cmd(host, f'msg * "{safe}"')

    def list_processes(self, host: str, limit: int = 50) -> CommandResult:
        limit = max(1, min(int(limit), 200))
        script = f"""
Get-Process -ErrorAction SilentlyContinue |
    Sort-Object WorkingSet64 -Descending |
    Select-Object -First {limit} Name,Id,CPU,WorkingSet,StartTime
"""
        return self.executor.execute_powershell_json(host, script)

    def process_status(
        self,
        host: str,
        *,
        pid: int | None = None,
        process_name: str | None = None,
    ) -> CommandResult:
        if pid is not None:
            process_id = int(pid)
            if process_id <= 0:
                return CommandResult.failure(
                    host,
                    "process_status",
                    "PID deve ser maior que zero.",
                )
            script = f"""
$p = Get-Process -Id {process_id} -ErrorAction SilentlyContinue
[pscustomobject]@{{
    Exists = [bool]$p
    Id = if ($p) {{ $p.Id }} else {{ $null }}
    Name = if ($p) {{ $p.Name }} else {{ $null }}
}}
"""
        elif process_name:
            safe_name = quote_powershell_literal(
                validate_process_name(process_name)
            )
            script = f"""
$p = @(Get-Process -Name {safe_name} -ErrorAction SilentlyContinue)
[pscustomobject]@{{
    Exists = [bool]($p.Count -gt 0)
    Count = $p.Count
    Ids = @($p | Select-Object -ExpandProperty Id)
    Name = {safe_name}
}}
"""
        else:
            return CommandResult.failure(
                host,
                "process_status",
                "Informe PID ou nome do processo.",
            )
        return self.executor.execute_powershell_json(host, script)

    def kill_process(self, host: str, process_name: str) -> CommandResult:
        safe_name = quote_powershell_literal(
            validate_process_name(process_name)
        )
        script = f"""
$items = @(Get-Process -Name {safe_name} -ErrorAction Stop)
$ids = @($items | Select-Object -ExpandProperty Id)
$items | Stop-Process -Force -ErrorAction Stop
[pscustomobject]@{{Stopped=$ids;Name={safe_name}}}
"""
        return self.executor.execute_mutating_powershell_json(host, script)

    def kill_process_pid(self, host: str, pid: int) -> CommandResult:
        process_id = int(pid)
        if process_id <= 0:
            return CommandResult.failure(
                host,
                "kill_process_pid",
                "PID deve ser maior que zero.",
            )
        script = f"""
$p = Get-Process -Id {process_id} -ErrorAction Stop
$name = $p.Name
Stop-Process -Id {process_id} -Force -ErrorAction Stop
[pscustomobject]@{{Stopped={process_id};Name=$name}}
"""
        return self.executor.execute_mutating_powershell_json(host, script)

    def list_services(self, host: str) -> CommandResult:
        script = r"""
Get-Service |
    Sort-Object Status,DisplayName |
    Select-Object Name,DisplayName,Status,StartType
"""
        return self.executor.execute_powershell_json(host, script)

    def service_status(
        self,
        host: str,
        service_name: str,
    ) -> CommandResult:
        safe = quote_powershell_literal(
            validate_safe_name(service_name, label="Serviço")
        )
        script = f"""
$svc = Get-Service -Name {safe} -ErrorAction Stop
[pscustomobject]@{{
    Name = $svc.Name
    DisplayName = $svc.DisplayName
    Status = $svc.Status.ToString()
    StartType = $svc.StartType.ToString()
}}
"""
        return self.executor.execute_powershell_json(host, script)

    def service_action(
        self,
        host: str,
        service_name: str,
        action: str,
    ) -> CommandResult:
        safe = quote_powershell_literal(
            validate_safe_name(service_name, label="Serviço")
        )
        actions = {
            "start": "Start-Service",
            "stop": "Stop-Service",
            "restart": "Restart-Service",
        }
        cmd = actions.get(action.casefold())
        if not cmd:
            return CommandResult.failure(
                host,
                action,
                "Ação de serviço inválida.",
            )
        script = f"""
{cmd} -Name {safe} -ErrorAction Stop
$svc = Get-Service -Name {safe} -ErrorAction Stop
[pscustomobject]@{{
    Name = $svc.Name
    Status = $svc.Status.ToString()
    StartType = $svc.StartType.ToString()
}}
"""
        return self.executor.execute_mutating_powershell_json(host, script)

    def set_service_startup(
        self,
        host: str,
        service_name: str,
        startup_type: str,
    ) -> CommandResult:
        safe = quote_powershell_literal(
            validate_safe_name(service_name, label="Serviço")
        )
        allowed = {
            "Automatic": "Automatic",
            "Manual": "Manual",
            "Disabled": "Disabled",
        }
        value = allowed.get(startup_type)
        if not value:
            return CommandResult.failure(
                host,
                "set_service_startup",
                "StartType inválido.",
            )
        script = f"""
Set-Service -Name {safe} -StartupType {value} -ErrorAction Stop
$svc = Get-Service -Name {safe} -ErrorAction Stop
[pscustomobject]@{{
    Name = $svc.Name
    Status = $svc.Status.ToString()
    StartType = $svc.StartType.ToString()
}}
"""
        return self.executor.execute_mutating_powershell_json(host, script)

    def gpupdate(self, host: str, force: bool = True) -> CommandResult:
        suffix = " /force" if force else ""
        return self.executor.execute_mutating_cmd(
            host,
            f"gpupdate{suffix}",
            timeout=180,
        )

    def logoff_session(self, host: str, session_id: int) -> CommandResult:
        value = int(session_id)
        if value < 0:
            return CommandResult.failure(
                host,
                "logoff_session",
                "ID de sessão inválido.",
            )
        return self.executor.execute_mutating_cmd(
            host,
            f"logoff {value}",
            timeout=60,
        )

    def restart(
        self,
        host: str,
        delay_seconds: int = 0,
        message: str = "Reinicialização administrativa",
    ) -> CommandResult:
        safe = message.replace('"', "'").replace("\r", " ").replace("\n", " ")
        return self.executor.execute_mutating_cmd(
            host,
            f'shutdown /r /t {max(0, int(delay_seconds))} /c "{safe}"',
        )

    def shutdown(
        self,
        host: str,
        delay_seconds: int = 0,
        message: str = "Desligamento administrativo",
    ) -> CommandResult:
        safe = message.replace('"', "'").replace("\r", " ").replace("\n", " ")
        return self.executor.execute_mutating_cmd(
            host,
            f'shutdown /s /t {max(0, int(delay_seconds))} /c "{safe}"',
        )

    def abort_shutdown(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(host, "shutdown /a")
