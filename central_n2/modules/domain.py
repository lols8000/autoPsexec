from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class DomainModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def status(self, host: str) -> CommandResult:
        script = r"""
$cs = Get-CimInstance Win32_ComputerSystem
$domain = $cs.Domain
$dc = $null
$secure = $null
try { $dc = (nltest /dsgetdc:$domain 2>$null | Out-String).Trim() } catch {}
try { $secure = Test-ComputerSecureChannel -ErrorAction Stop } catch {}
$time = $null
try { $time = (w32tm /query /status 2>$null | Out-String).Trim() } catch {}
[pscustomobject]@{
    Domain = $domain
    PartOfDomain = $cs.PartOfDomain
    SecureChannel = $secure
    DomainController = $dc
    TimeStatus = $time
}
"""
        return self.executor.execute_powershell_json(host, script, timeout=90)

    def gpresult(self, host: str) -> CommandResult:
        return self.executor.execute_cmd(
            host,
            "gpresult /r /scope computer",
            timeout=120,
        )

    def gpupdate(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(
            host,
            "gpupdate /force",
            timeout=300,
        )

    def repair_secure_channel(self, host: str) -> CommandResult:
        script = r"""
$before = Test-ComputerSecureChannel -ErrorAction SilentlyContinue
$repair = Test-ComputerSecureChannel -Repair -ErrorAction Stop
$after = Test-ComputerSecureChannel -ErrorAction Stop
[pscustomobject]@{
    Before = $before
    RepairResult = $repair
    SecureChannel = $after
}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=180,
        )

    def restart_time_service(self, host: str) -> CommandResult:
        script = r"""
Restart-Service w32time -Force -ErrorAction Stop
$svc = Get-Service w32time -ErrorAction Stop
[pscustomobject]@{
    Name = $svc.Name
    Status = $svc.Status.ToString()
    StartType = $svc.StartType.ToString()
}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=90,
        )

    def resync_time(self, host: str) -> CommandResult:
        script = r"""
Restart-Service w32time -Force -ErrorAction Stop
$output = & w32tm /resync /rediscover 2>&1 | Out-String
$code = $LASTEXITCODE
if ($code -ne 0) {
    throw "w32tm /resync retornou exit code $code. $output"
}
$status = & w32tm /query /status 2>&1 | Out-String
[pscustomobject]@{
    ResyncSucceeded = $true
    Output = $output.Trim()
    Status = $status.Trim()
}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=120,
        )

    def purge_system_kerberos(self, host: str) -> CommandResult:
        script = r"""
$output = & klist.exe -li 0x3e7 purge 2>&1 | Out-String
$code = $LASTEXITCODE
if ($code -ne 0) {
    throw "klist purge retornou exit code $code. $output"
}
[pscustomobject]@{
    Purged = $true
    LogonId = '0x3e7'
    Output = $output.Trim()
}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=90,
        )
