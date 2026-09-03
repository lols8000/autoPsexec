from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class TasksModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def list_tasks(self, host: str) -> CommandResult:
        script = r'''
Get-ScheduledTask -ErrorAction SilentlyContinue | ForEach-Object {
  $info = $_ | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
  [pscustomobject]@{
    TaskName=$_.TaskName
    TaskPath=$_.TaskPath
    State=$_.State
    LastRunTime=$info.LastRunTime
    LastTaskResult=$info.LastTaskResult
    NextRunTime=$info.NextRunTime
  }
} | Sort-Object LastTaskResult,TaskPath,TaskName
'''
        return self.executor.execute_powershell_json(host, script, timeout=180)

    def failed_tasks(self, host: str) -> CommandResult:
        script = r'''
Get-ScheduledTask -ErrorAction SilentlyContinue | ForEach-Object {
  $info = $_ | Get-ScheduledTaskInfo -ErrorAction SilentlyContinue
  if($info -and $info.LastTaskResult -ne 0){
    [pscustomobject]@{TaskName=$_.TaskName;TaskPath=$_.TaskPath;State=$_.State;LastRunTime=$info.LastRunTime;LastTaskResult=$info.LastTaskResult}
  }
} | Sort-Object LastRunTime -Descending
'''
        return self.executor.execute_powershell_json(host, script, timeout=180)
