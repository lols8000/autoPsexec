from __future__ import annotations
from datetime import datetime
class SupportReportBuilder:
    def build(self,*,host,user=None,problem=None,diagnosis=None,actions=None,validation=None,result=None):
        return {"generated_at":datetime.now().astimezone().isoformat(timespec="seconds"),"host":host,"user":user,"problem":problem,"diagnosis":diagnosis,"actions":actions or [],"validation":validation,"result":result}
