from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult


class DomainModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def status(self, host: str) -> CommandResult:
        script = r'''
$cs = Get-CimInstance Win32_ComputerSystem
$domain = $cs.Domain
$dc = $null
$secure = $null
try { $dc = (nltest /dsgetdc:$domain 2>$null | Out-String).Trim() } catch {}
try { $secure = Test-ComputerSecureChannel -ErrorAction Stop } catch {}
$time = $null
try { $time = (w32tm /query /status 2>$null | Out-String).Trim() } catch {}
[pscustomobject]@{
 Domain=$domain
 PartOfDomain=$cs.PartOfDomain
 SecureChannel=$secure
 DomainController=$dc
 TimeStatus=$time
}
'''
        return self.executor.execute_powershell_json(host, script, timeout=90)

    def gpresult(self, host: str) -> CommandResult:
        return self.executor.execute_cmd(host, "gpresult /r /scope computer", timeout=120)

    def repair_secure_channel(self, host: str) -> CommandResult:
        script = "Test-ComputerSecureChannel -Repair -ErrorAction Stop"
        return self.executor.execute_remote_powershell_with_fallback(host, script, timeout=120)
