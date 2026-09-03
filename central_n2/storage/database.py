from __future__ import annotations

import json
import sqlite3
from dataclasses import asdict, is_dataclass
from datetime import datetime, timedelta, timezone
from enum import Enum
from pathlib import Path
from typing import Any

from .diff import diff_values


class CentralDatabase:
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path, timeout=10)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA foreign_keys=ON")
        return connection

    @staticmethod
    def _jsonable(value: Any) -> Any:
        if is_dataclass(value):
            value = asdict(value)
        if isinstance(value, Enum):
            return value.value
        if isinstance(value, dict):
            return {str(k): CentralDatabase._jsonable(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [CentralDatabase._jsonable(v) for v in value]
        return value

    @classmethod
    def _encode(cls, payload: Any) -> str:
        return json.dumps(
            cls._jsonable(payload),
            ensure_ascii=False,
            default=str,
        )

    def initialize(self) -> None:
        schema = """
        CREATE TABLE IF NOT EXISTS hosts(
            host TEXT PRIMARY KEY,
            last_seen TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS snapshots(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            host TEXT NOT NULL,
            created_at TEXT NOT NULL,
            kind TEXT NOT NULL,
            payload TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_snapshots_host_created
            ON snapshots(host, created_at DESC);

        CREATE TABLE IF NOT EXISTS jobs(
            job_id TEXT PRIMARY KEY,
            host TEXT NOT NULL,
            created_at TEXT NOT NULL,
            state TEXT NOT NULL,
            label TEXT NOT NULL,
            payload TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_jobs_host_created
            ON jobs(host, created_at DESC);

        CREATE TABLE IF NOT EXISTS findings(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            host TEXT NOT NULL,
            created_at TEXT NOT NULL,
            finding_id TEXT NOT NULL,
            severity TEXT NOT NULL,
            payload TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS remediations(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            host TEXT NOT NULL,
            created_at TEXT NOT NULL,
            action TEXT NOT NULL,
            success INTEGER NOT NULL,
            payload TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS reports(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            host TEXT NOT NULL,
            created_at TEXT NOT NULL,
            format TEXT NOT NULL,
            path TEXT,
            payload TEXT NOT NULL
        );
        """
        with self._connect() as con:
            con.executescript(schema)

    def save_snapshot(self, host: str, payload: Any, *, kind: str = "health") -> int:
        now = datetime.now(timezone.utc).isoformat()
        encoded = self._encode(payload)
        with self._connect() as con:
            con.execute(
                """
                INSERT INTO hosts(host,last_seen)
                VALUES(?,?)
                ON CONFLICT(host) DO UPDATE SET last_seen=excluded.last_seen
                """,
                (host, now),
            )
            cursor = con.execute(
                "INSERT INTO snapshots(host,created_at,kind,payload) VALUES(?,?,?,?)",
                (host, now, kind, encoded),
            )
            return int(cursor.lastrowid)

    def recent_snapshots(
        self,
        host: str,
        *,
        kind: str | None = None,
        limit: int = 20,
    ) -> list[dict[str, Any]]:
        sql = (
            "SELECT id,host,created_at,kind,payload "
            "FROM snapshots WHERE lower(host)=lower(?)"
        )
        args: list[Any] = [host]
        if kind:
            sql += " AND kind=?"
            args.append(kind)
        sql += " ORDER BY id DESC LIMIT ?"
        args.append(max(1, int(limit)))

        with self._connect() as con:
            rows = con.execute(sql, args).fetchall()

        return [
            {**dict(row), "payload": json.loads(row["payload"])}
            for row in rows
        ]

    def diff_latest(self, host: str, *, kind: str = "health") -> list[dict[str, Any]]:
        rows = self.recent_snapshots(host, kind=kind, limit=2)
        if len(rows) < 2:
            return []
        return diff_values(rows[1]["payload"], rows[0]["payload"])

    def save_job(self, record: Any) -> None:
        payload = self._jsonable(record)
        job_id = str(payload["job_id"])
        host = str(payload["host"])
        created_at = str(payload["created_at"])
        state = str(payload["state"])
        label = str(payload["label"])
        with self._connect() as con:
            con.execute(
                """
                INSERT INTO jobs(job_id,host,created_at,state,label,payload)
                VALUES(?,?,?,?,?,?)
                ON CONFLICT(job_id) DO UPDATE SET
                    state=excluded.state,
                    label=excluded.label,
                    payload=excluded.payload
                """,
                (
                    job_id,
                    host,
                    created_at,
                    state,
                    label,
                    self._encode(payload),
                ),
            )

    def recent_jobs(self, host: str | None = None, *, limit: int = 50) -> list[dict[str, Any]]:
        if host:
            sql = (
                "SELECT job_id,host,created_at,state,label,payload "
                "FROM jobs WHERE lower(host)=lower(?) "
                "ORDER BY created_at DESC LIMIT ?"
            )
            args: tuple[Any, ...] = (host, max(1, int(limit)))
        else:
            sql = (
                "SELECT job_id,host,created_at,state,label,payload "
                "FROM jobs ORDER BY created_at DESC LIMIT ?"
            )
            args = (max(1, int(limit)),)

        with self._connect() as con:
            rows = con.execute(sql, args).fetchall()
        return [
            {**dict(row), "payload": json.loads(row["payload"] or "{}")}
            for row in rows
        ]

    def save_finding(
        self,
        host: str,
        finding_id: str,
        severity: str,
        payload: Any,
    ) -> None:
        with self._connect() as con:
            con.execute(
                """
                INSERT INTO findings(host,created_at,finding_id,severity,payload)
                VALUES(?,?,?,?,?)
                """,
                (
                    host,
                    datetime.now(timezone.utc).isoformat(),
                    finding_id,
                    severity,
                    self._encode(payload),
                ),
            )

    def save_remediation(
        self,
        host: str,
        action: str,
        success: bool,
        payload: Any,
    ) -> None:
        with self._connect() as con:
            con.execute(
                """
                INSERT INTO remediations(host,created_at,action,success,payload)
                VALUES(?,?,?,?,?)
                """,
                (
                    host,
                    datetime.now(timezone.utc).isoformat(),
                    action,
                    int(success),
                    self._encode(payload),
                ),
            )

    def recent_remediations(self, host: str, *, limit: int = 20) -> list[dict[str, Any]]:
        with self._connect() as con:
            rows = con.execute(
                """
                SELECT id,host,created_at,action,success,payload
                FROM remediations
                WHERE lower(host)=lower(?)
                ORDER BY id DESC
                LIMIT ?
                """,
                (host, max(1, int(limit))),
            ).fetchall()
        return [
            {**dict(row), "payload": json.loads(row["payload"])}
            for row in rows
        ]

    def save_report(
        self,
        host: str,
        fmt: str,
        payload: str,
        *,
        path: str | None = None,
    ) -> None:
        with self._connect() as con:
            con.execute(
                """
                INSERT INTO reports(host,created_at,format,path,payload)
                VALUES(?,?,?,?,?)
                """,
                (
                    host,
                    datetime.now(timezone.utc).isoformat(),
                    fmt,
                    path,
                    payload,
                ),
            )

    def prune(self, retention_days: int) -> int:
        cutoff = (
            datetime.now(timezone.utc)
            - timedelta(days=max(1, retention_days))
        ).isoformat()
        with self._connect() as con:
            cursor = con.execute(
                "DELETE FROM snapshots WHERE created_at < ?",
                (cutoff,),
            )
            return cursor.rowcount
