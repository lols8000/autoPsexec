from __future__ import annotations
import json,urllib.parse,urllib.request
from typing import Any
class GLPIError(RuntimeError): pass
class GLPIClient:
    def __init__(self,base_url,app_token,user_token,*,timeout=20):
        self.base_url=base_url.rstrip("/"); self.app_token=app_token; self.user_token=user_token; self.timeout=timeout; self.session_token=None
    def _headers(self,session=True):
        h={"Content-Type":"application/json","App-Token":self.app_token}
        if session and self.session_token: h["Session-Token"]=self.session_token
        elif not session: h["Authorization"]=f"user_token {self.user_token}"
        return h
    def _request(self,method,path,*,payload=None,session=True)->Any:
        if not self.base_url: raise GLPIError("glpi_api.base_url não configurado.")
        data=None if payload is None else json.dumps(payload).encode("utf-8")
        req=urllib.request.Request(f"{self.base_url}/{path.lstrip('/')}",data=data,method=method,headers=self._headers(session))
        try:
            with urllib.request.urlopen(req,timeout=self.timeout) as res: raw=res.read().decode("utf-8")
        except Exception as exc: raise GLPIError(str(exc)) from exc
        return json.loads(raw) if raw else None
    def init_session(self):
        data=self._request("GET","initSession",session=False); token=(data or {}).get("session_token")
        if not token: raise GLPIError("GLPI não retornou session_token.")
        self.session_token=token; return token
    def kill_session(self):
        if self.session_token:
            try: self._request("GET","killSession")
            finally: self.session_token=None
    def search_computer(self,name):
        if not self.session_token: self.init_session()
        q=urllib.parse.urlencode({"criteria[0][field]":1,"criteria[0][searchtype]":"contains","criteria[0][value]":name})
        return self._request("GET",f"search/Computer?{q}")
    def add_ticket_followup(self,ticket_id,content):
        if not self.session_token: self.init_session()
        return self._request("POST","TicketFollowup",payload={"input":{"items_id":int(ticket_id),"content":content}})
