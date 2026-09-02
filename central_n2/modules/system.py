from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class SystemModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def sessions(self, host: str) -> CommandResult:
        return self.executor.execute_cmd(host, "quser")

    def send_message(self, host: str, message: str) -> CommandResult:
        safe = message.replace('"', "'").replace("\r", " ").replace("\n", " ")
        return self.executor.execute_cmd(host, f'msg * "{safe}"')

    def list_processes(self, host: str, limit: int = 50) -> CommandResult:
        script = f'''
Get-Process | Sort-Object CPU -Descending | Select-Object -First {int(limit)} Name,Id,CPU,WorkingSet,StartTime
'''
        return self.executor.execute_powershell_json(host, script)

    def kill_process(self, host: str, process_name: str) -> CommandResult:
        safe = process_name.replace("'", "''")
        return self.executor.execute_remote_powershell_with_fallback(
            host,
            f"Stop-Process -Name '{safe}' -Force -ErrorAction Stop",
        )

    def list_services(self, host: str) -> CommandResult:
        script = r'''
Get-Service | Sort-Object Status,DisplayName | Select-Object Name,DisplayName,Status,StartType
'''
        return self.executor.execute_powershell_json(host, script)

    def service_action(self, host: str, service_name: str, action: str) -> CommandResult:
        safe = service_name.replace("'", "''")
        actions = {
            "start": "Start-Service",
            "stop": "Stop-Service",
            "restart": "Restart-Service",
        }
        cmd = actions.get(action.lower())
        if not cmd:
            return CommandResult.failure(host, action, "Ação de serviço inválida.")
        return self.executor.execute_remote_powershell_with_fallback(
            host,
            f"{cmd} -Name '{safe}' -ErrorAction Stop",
        )

    def gpupdate(self, host: str, force: bool = True) -> CommandResult:
        suffix = " /force" if force else ""
        return self.executor.execute_cmd(host, f"gpupdate{suffix}", timeout=180)

    def restart(self, host: str, delay_seconds: int = 0, message: str = "Reinicialização administrativa") -> CommandResult:
        safe = message.replace('"', "'")
        return self.executor.execute_cmd(host, f'shutdown /r /t {int(delay_seconds)} /c "{safe}"')

    def shutdown(self, host: str, delay_seconds: int = 0, message: str = "Desligamento administrativo") -> CommandResult:
        safe = message.replace('"', "'")
        return self.executor.execute_cmd(host, f'shutdown /s /t {int(delay_seconds)} /c "{safe}"')

    def abort_shutdown(self, host: str) -> CommandResult:
        return self.executor.execute_cmd(host, "shutdown /a")
