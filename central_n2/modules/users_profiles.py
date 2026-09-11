from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import validate_sid


class UsersProfilesModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def local_admins(self, host: str) -> CommandResult:
        script = r"""
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
        script = r"""
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

    def profile_status(self, host: str, sid: str) -> CommandResult:
        safe_sid = validate_sid(sid)
        script = f"""
$p = Get-CimInstance Win32_UserProfile -Filter "SID='{safe_sid}'" -ErrorAction SilentlyContinue
[pscustomobject]@{{
    Exists = [bool]$p
    SID = '{safe_sid}'
    LocalPath = if ($p) {{ $p.LocalPath }} else {{ $null }}
    Loaded = if ($p) {{ [bool]$p.Loaded }} else {{ $null }}
    Special = if ($p) {{ [bool]$p.Special }} else {{ $null }}
    LastUseTime = if ($p) {{ $p.LastUseTime }} else {{ $null }}
}}
"""
        return self.executor.execute_powershell_json(host, script)

    def clean_profile_temp(self, host: str, sid: str) -> CommandResult:
        safe_sid = validate_sid(sid)
        script = fr"""
$p = Get-CimInstance Win32_UserProfile -Filter "SID='{safe_sid}'" -ErrorAction Stop
if (-not $p) {{ throw 'Perfil não encontrado.' }}
if ($p.Special) {{ throw 'Perfil especial não pode ser limpo por esta ação.' }}
$target = Join-Path $p.LocalPath 'AppData\Local\Temp'
$before = 0
$after = 0
if (Test-Path $target) {{
    $value = (
        Get-ChildItem $target -Force -Recurse -File -ErrorAction SilentlyContinue |
        Measure-Object Length -Sum
    ).Sum
    if ($value) {{ $before = $value }}
    Get-ChildItem $target -Force -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    $remaining = (
        Get-ChildItem $target -Force -Recurse -File -ErrorAction SilentlyContinue |
        Measure-Object Length -Sum
    ).Sum
    if ($remaining) {{ $after = $remaining }}
}}
[pscustomobject]@{{
    SID = '{safe_sid}'
    Path = $target
    BeforeGB = [math]::Round($before / 1GB, 2)
    AfterGB = [math]::Round($after / 1GB, 2)
    RecoveredGB = [math]::Round(($before - $after) / 1GB, 2)
}}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=300,
        )

    def remove_profile(self, host: str, sid: str) -> CommandResult:
        safe_sid = validate_sid(sid)
        script = f"""
$p = Get-CimInstance Win32_UserProfile -Filter "SID='{safe_sid}'" -ErrorAction Stop
if (-not $p) {{ throw 'Perfil não encontrado.' }}
if ($p.Loaded) {{ throw 'Perfil carregado não pode ser removido.' }}
if ($p.Special) {{ throw 'Perfil especial não pode ser removido.' }}
$path = $p.LocalPath
Remove-CimInstance -InputObject $p -ErrorAction Stop
$remaining = Get-CimInstance Win32_UserProfile -Filter "SID='{safe_sid}'" -ErrorAction SilentlyContinue
[pscustomobject]@{{
    SID = '{safe_sid}'
    Path = $path
    Removed = [bool](-not $remaining)
}}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=180,
        )
