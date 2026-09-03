from __future__ import annotations
import re
from dataclasses import dataclass
_PERCENT=re.compile(r"(?<!\d)(\d{1,3}(?:[.,]\d+)?)\s*%")
@dataclass(slots=True)
class ProgressEvent:percent:float|None;message:str
class ProgressParser:
    @staticmethod
    def parse(line:str)->ProgressEvent:
        m=_PERCENT.search(line);value=None
        if m:value=max(0.0,min(100.0,float(m.group(1).replace(",","."))))
        return ProgressEvent(value,line.strip())
