from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import validate_sid


class UsersProfilesModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def local_admins(self, host: str) -> CommandResult:
        script = """
try {
    Get-LocalGroupMember -Group 'Administrators' -ErrorAction Stop |
        Select-Object Name,ObjectClass,PrincipalSource
} catch {
    Get-LocalGroupMember -Group 'Administradores' -ErrorAction Stop |
        Select-Object Name,ObjectClass,PrincipalSource
}
"""
        return self.executor.execute_powershell_json(host, script)

    def profiles(self, host: str) -> CommandResult:
        script = """
$users = Get-CimInstance Win32_UserProfile | Where-Object { -not $_.Special }
foreach ($u in $users) {
    $path = $u.LocalPath
    $size = $null
    try {
        $bytes = (
            Get-ChildItem -LiteralPath $path -Force -Recurse -File -ErrorAction SilentlyContinue |
            Measure-Object Length -Sum
        ).Sum
        if ($bytes -ne $null) {
            $size = [math]::Round($bytes / 1GB, 2)
        }
    } catch {}
    [pscustomobject]@{
        LocalPath = $path
        SID = $u.SID
        Loaded = $u.Loaded
        LastUseTime = $u.LastUseTime
        SizeGB = $size
    }
}
"""
        result = self.executor.execute_powershell_json(host, script, timeout=600)
        result.metadata["heavy_read"] = True
        return result

    def remove_profile(self, host: str, sid: str) -> CommandResult:
        safe_sid = validate_sid(sid)
        script = (
            "Get-CimInstance Win32_UserProfile "
            f"-Filter \"SID='{safe_sid}'\" | "
            "Where-Object { -not $_.Loaded -and -not $_.Special } | "
            "Remove-CimInstance -ErrorAction Stop"
        )
        return self.executor.execute_mutating_powershell(
            host,
            script,
            timeout=120,
        )
