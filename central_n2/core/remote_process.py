from __future__ import annotations
import base64
from core.result import CommandResult
class RemoteProcessController:
    def __init__(self,executor)->None:self.executor=executor
    def start_powershell(self,host:str,script:str)->CommandResult:
        encoded=base64.b64encode(script.encode("utf-16le")).decode("ascii")
        ps=f"$p=Start-Process powershell.exe -ArgumentList '-NoProfile','-EncodedCommand','{encoded}' -PassThru -WindowStyle Hidden;[pscustomobject]@{{PID=$p.Id}}"
        return self.executor.execute_powershell_json(host,ps,timeout=60)
    def cancel(self,host:str,pid:int)->CommandResult:
        return self.executor.execute_remote_powershell_with_fallback(host,f"Stop-Process -Id {int(pid)} -Force -ErrorAction Stop",timeout=60)
