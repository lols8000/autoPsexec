from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class UsersProfilesModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def local_admins(self, host: str) -> CommandResult:
        script = r'''
try {
  Get-LocalGroupMember -Group 'Administrators' -ErrorAction Stop |
    Select-Object Name,ObjectClass,PrincipalSource
} catch {
  try {
    Get-LocalGroupMember -Group 'Administradores' -ErrorAction Stop |
      Select-Object Name,ObjectClass,PrincipalSource
  } catch { throw }
}
'''
        return self.executor.execute_powershell_json(host, script)

    def profiles(self, host: str) -> CommandResult:
        script = r'''
$users = Get-CimInstance Win32_UserProfile | Where-Object { -not $_.Special }
foreach($u in $users){
  $path = $u.LocalPath
  $size = $null
  try {
    $bytes = (Get-ChildItem -LiteralPath $path -Force -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
    if($bytes -ne $null){$size=[math]::Round($bytes/1GB,2)}
  } catch {}
  [pscustomobject]@{
    LocalPath=$path
    SID=$u.SID
    Loaded=$u.Loaded
    LastUseTime=$u.LastUseTime
    SizeGB=$size
  }
}
'''
        return self.executor.execute_powershell_json(host, script, timeout=180)

    def remove_profile(self, host: str, sid: str) -> CommandResult:
        safe = sid.replace("'", "''")
        script = f"Get-CimInstance Win32_UserProfile -Filter \"SID='{safe}'\" | Remove-CimInstance -ErrorAction Stop"
        return self.executor.execute_remote_powershell_with_fallback(host, script, timeout=120)
