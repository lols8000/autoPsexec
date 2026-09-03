from __future__ import annotations
import json,urllib.request
from dataclasses import dataclass
from pathlib import Path
@dataclass(slots=True)
class ReleaseInfo:
    current:str;latest:str;update_available:bool;html_url:str|None;assets:list[dict]
class UpdateManager:
    def __init__(self,repository:str,current_version:str,*,timeout:int=10)->None:
        self.repository=repository;self.current_version=current_version.lstrip("v");self.timeout=timeout
    @staticmethod
    def _version_tuple(value:str):
        parts=[]
        for token in value.lstrip("v").split("."):
            digits="".join(ch for ch in token if ch.isdigit());parts.append(int(digits or 0))
        return tuple(parts)
    def check_latest(self)->ReleaseInfo:
        req=urllib.request.Request(f"https://api.github.com/repos/{self.repository}/releases/latest",headers={"Accept":"application/vnd.github+json","User-Agent":"CentralN2"})
        with urllib.request.urlopen(req,timeout=self.timeout) as res:data=json.loads(res.read().decode("utf-8"))
        latest=str(data.get("tag_name") or "0.0.0").lstrip("v")
        return ReleaseInfo(self.current_version,latest,self._version_tuple(latest)>self._version_tuple(self.current_version),data.get("html_url"),data.get("assets") or [])
    def download_asset(self,asset:dict,directory:str|Path)->Path:
        d=Path(directory);d.mkdir(parents=True,exist_ok=True);target=d/(asset.get("name") or "update.bin");url=asset.get("browser_download_url")
        if not url:raise ValueError("Asset sem browser_download_url")
        req=urllib.request.Request(url,headers={"User-Agent":"CentralN2"})
        with urllib.request.urlopen(req,timeout=60) as res,target.open("wb") as fh:
            while True:
                chunk=res.read(1024*1024)
                if not chunk:break
                fh.write(chunk)
        return target
