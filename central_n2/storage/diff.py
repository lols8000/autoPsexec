from __future__ import annotations
from typing import Any
def diff_values(before:Any,after:Any,path:str="")->list[dict[str,Any]]:
    out=[]
    if isinstance(before,dict) and isinstance(after,dict):
        for key in sorted(set(before)|set(after)):
            child=f"{path}.{key}" if path else str(key)
            if key not in before: out.append({"path":child,"type":"added","before":None,"after":after[key]})
            elif key not in after: out.append({"path":child,"type":"removed","before":before[key],"after":None})
            else: out.extend(diff_values(before[key],after[key],child))
        return out
    if before!=after: out.append({"path":path or "$","type":"changed","before":before,"after":after})
    return out
