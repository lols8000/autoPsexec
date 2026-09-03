from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime
from typing import Any,Callable
from core.result import CommandResult

@dataclass(frozen=True,slots=True)
class RemediationSpec:
    key:str; title:str; impact:str; requires_confirmation:bool=True; requires_reboot:bool=False; disruptive:bool=False; rollback:str|None=None
@dataclass(slots=True)
class RemediationResult:
    spec:RemediationSpec; host:str; before:Any; command_result:CommandResult; after:Any; started_at:str; finished_at:str

class RemediationEngine:
    def execute(self,host,spec,action:Callable[[str],CommandResult],*,snapshotter=None):
        started=datetime.now().astimezone().isoformat(timespec="seconds")
        before=snapshotter(host) if snapshotter else None
        result=action(host)
        after=snapshotter(host) if snapshotter and result.success else None
        return RemediationResult(spec,host,before,result,after,started,datetime.now().astimezone().isoformat(timespec="seconds"))
