from __future__ import annotations
import json,sqlite3
from datetime import datetime,timedelta,timezone
from pathlib import Path
from typing import Any
from .diff import diff_values

class CentralDatabase:
    def __init__(self,path:str|Path)->None:
        self.path=Path(path); self.path.parent.mkdir(parents=True,exist_ok=True); self.initialize()
    def _connect(self):
        con=sqlite3.connect(self.path,timeout=10); con.row_factory=sqlite3.Row; con.execute("PRAGMA journal_mode=WAL"); con.execute("PRAGMA foreign_keys=ON"); return con
    def initialize(self):
        schema="""
        CREATE TABLE IF NOT EXISTS hosts(host TEXT PRIMARY KEY,last_seen TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS snapshots(id INTEGER PRIMARY KEY AUTOINCREMENT,host TEXT NOT NULL,created_at TEXT NOT NULL,kind TEXT NOT NULL,payload TEXT NOT NULL);
        CREATE INDEX IF NOT EXISTS idx_snapshots_host_created ON snapshots(host,created_at DESC);
        CREATE TABLE IF NOT EXISTS jobs(job_id TEXT PRIMARY KEY,host TEXT NOT NULL,created_at TEXT NOT NULL,state TEXT NOT NULL,label TEXT NOT NULL,payload TEXT);
        CREATE TABLE IF NOT EXISTS findings(id INTEGER PRIMARY KEY AUTOINCREMENT,host TEXT NOT NULL,created_at TEXT NOT NULL,finding_id TEXT NOT NULL,severity TEXT NOT NULL,payload TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS remediations(id INTEGER PRIMARY KEY AUTOINCREMENT,host TEXT NOT NULL,created_at TEXT NOT NULL,action TEXT NOT NULL,success INTEGER NOT NULL,payload TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS reports(id INTEGER PRIMARY KEY AUTOINCREMENT,host TEXT NOT NULL,created_at TEXT NOT NULL,format TEXT NOT NULL,path TEXT,payload TEXT NOT NULL);
        """
        with self._connect() as con: con.executescript(schema)
    def save_snapshot(self,host,payload,*,kind="health"):
        now=datetime.now(timezone.utc).isoformat(); enc=json.dumps(payload,ensure_ascii=False,default=str)
        with self._connect() as con:
            con.execute("INSERT INTO hosts(host,last_seen) VALUES(?,?) ON CONFLICT(host) DO UPDATE SET last_seen=excluded.last_seen",(host,now))
            return int(con.execute("INSERT INTO snapshots(host,created_at,kind,payload) VALUES(?,?,?,?)",(host,now,kind,enc)).lastrowid)
    def recent_snapshots(self,host,*,kind=None,limit=20):
        sql="SELECT id,host,created_at,kind,payload FROM snapshots WHERE lower(host)=lower(?)"; args=[host]
        if kind: sql+=" AND kind=?"; args.append(kind)
        sql+=" ORDER BY id DESC LIMIT ?"; args.append(max(1,int(limit)))
        with self._connect() as con: rows=con.execute(sql,args).fetchall()
        return [{**dict(r),"payload":json.loads(r["payload"])} for r in rows]
    def diff_latest(self,host,*,kind="health"):
        rows=self.recent_snapshots(host,kind=kind,limit=2)
        return [] if len(rows)<2 else diff_values(rows[1]["payload"],rows[0]["payload"])
    def save_finding(self,host,finding_id,severity,payload):
        with self._connect() as con: con.execute("INSERT INTO findings(host,created_at,finding_id,severity,payload) VALUES(?,?,?,?,?)",(host,datetime.now(timezone.utc).isoformat(),finding_id,severity,json.dumps(payload,ensure_ascii=False,default=str)))
    def save_remediation(self,host,action,success,payload):
        with self._connect() as con: con.execute("INSERT INTO remediations(host,created_at,action,success,payload) VALUES(?,?,?,?,?)",(host,datetime.now(timezone.utc).isoformat(),action,int(success),json.dumps(payload,ensure_ascii=False,default=str)))
    def save_report(self,host,fmt,payload,*,path=None):
        with self._connect() as con: con.execute("INSERT INTO reports(host,created_at,format,path,payload) VALUES(?,?,?,?,?)",(host,datetime.now(timezone.utc).isoformat(),fmt,path,payload))
    def prune(self,retention_days):
        cutoff=(datetime.now(timezone.utc)-timedelta(days=max(1,retention_days))).isoformat()
        with self._connect() as con: return con.execute("DELETE FROM snapshots WHERE created_at < ?",(cutoff,)).rowcount
