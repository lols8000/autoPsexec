from __future__ import annotations

import json
import sqlite3
from dataclasses import asdict, is_dataclass
from datetime import datetime, timedelta, timezone
from enum import Enum
from pathlib import Path
from typing import Any

from core.redaction import redact

from .diff import diff_values


class CentralDatabase:
    SCHEMA_VERSION = 4

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(
            self.path,
            timeout=10,
            isolation_level="DEFERRED",
        )
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA foreign_keys=ON")
        connection.execute("PRAGMA busy_timeout=10000")
        return connection

    @staticmethod
    def _jsonable(value: Any) -> Any:
        if is_dataclass(value):
            value = asdict(value)
        if isinstance(value, Enum):
            return value.value
        if isinstance(value, dict):
            return {
                str(key): CentralDatabase._jsonable(item)
                for key, item in value.items()
            }
        if isinstance(value, (list, tuple)):
            return [
                CentralDatabase._jsonable(item)
                for item in value
            ]
        return value

    @classmethod
    def _encode(cls, payload: Any) -> str:
        return json.dumps(
            cls._jsonable(payload),
            ensure_ascii=False,
            default=str,
        )

    @staticmethod
    def _column_exists(
        connection: sqlite3.Connection,
        table: str,
        column: str,
    ) -> bool:
        rows = connection.execute(
            f"PRAGMA table_info({table})"
        ).fetchall()
        return any(row["name"] == column for row in rows)

    def initialize(self) -> None:
        with self._connect() as connection:
            version = int(
                connection.execute("PRAGMA user_version").fetchone()[0]
            )
            if version < 1:
                self._migration_1(connection)
                version = 1
            if version < 2:
                self._migration_2(connection)
                version = 2
            if version < 3:
                self._migration_3(connection)
                version = 3
            if version < 4:
                self._migration_4(connection)
                version = 4
            connection.execute(f"PRAGMA user_version={version}")

            if version != self.SCHEMA_VERSION:
                raise RuntimeError(
                    f"Schema SQLite inesperado: {version}; "
                    f"esperado {self.SCHEMA_VERSION}."
                )

    @staticmethod
    def _migration_1(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
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
        )

    @classmethod
    def _migration_2(cls, connection: sqlite3.Connection) -> None:
        for table in (
            "snapshots",
            "jobs",
            "findings",
            "remediations",
            "reports",
        ):
            if not cls._column_exists(
                connection,
                table,
                "correlation_id",
            ):
                connection.execute(
                    f"ALTER TABLE {table} "
                    "ADD COLUMN correlation_id TEXT"
                )

        connection.executescript(
            """
            CREATE INDEX IF NOT EXISTS idx_snapshots_correlation
                ON snapshots(correlation_id);
            CREATE INDEX IF NOT EXISTS idx_jobs_correlation
                ON jobs(correlation_id);
            CREATE INDEX IF NOT EXISTS idx_findings_host_created
                ON findings(host, created_at DESC);
            CREATE INDEX IF NOT EXISTS idx_findings_correlation
                ON findings(correlation_id);
            CREATE INDEX IF NOT EXISTS idx_remediations_host_created
                ON remediations(host, created_at DESC);
            CREATE INDEX IF NOT EXISTS idx_remediations_correlation
                ON remediations(correlation_id);
            CREATE INDEX IF NOT EXISTS idx_reports_host_created
                ON reports(host, created_at DESC);
            CREATE INDEX IF NOT EXISTS idx_reports_correlation
                ON reports(correlation_id);
            """
        )

    @staticmethod
    def _migration_3(connection: sqlite3.Connection) -> None:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS executions(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                host TEXT NOT NULL,
                created_at TEXT NOT NULL,
                action TEXT NOT NULL,
                validation_state TEXT NOT NULL,
                payload TEXT NOT NULL,
                correlation_id TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_executions_host_created
                ON executions(host, created_at DESC);
            CREATE INDEX IF NOT EXISTS idx_executions_correlation
                ON executions(correlation_id);
            CREATE INDEX IF NOT EXISTS idx_executions_action
                ON executions(action);
            """
        )

    @classmethod
    def _migration_4(cls, connection: sqlite3.Connection) -> None:
        columns = {
            "operator": "TEXT",
            "action_version": "INTEGER",
            "transport": "TEXT",
            "started_at": "TEXT",
            "finished_at": "TEXT",
            "duration_ms": "INTEGER",
            "risk": "TEXT",
            "parameters": "TEXT",
            "rollback_available": "INTEGER",
            "rollback_of": "INTEGER",
            "is_rollback": "INTEGER DEFAULT 0",
        }
        for name, definition in columns.items():
            if not cls._column_exists(
                connection,
                "executions",
                name,
            ):
                connection.execute(
                    f"ALTER TABLE executions "
                    f"ADD COLUMN {name} {definition}"
                )

        connection.executescript(
            """
            CREATE INDEX IF NOT EXISTS idx_executions_operator
                ON executions(operator);
            CREATE INDEX IF NOT EXISTS idx_executions_transport
                ON executions(transport);
            CREATE INDEX IF NOT EXISTS idx_executions_rollback_of
                ON executions(rollback_of);
            """
        )

    def integrity_check(self) -> bool:
        with self._connect() as connection:
            value = connection.execute(
                "PRAGMA quick_check"
            ).fetchone()[0]
        return str(value).casefold() == "ok"

    def _touch_host(
        self,
        connection: sqlite3.Connection,
        host: str,
        now: str,
    ) -> None:
        connection.execute(
            """
            INSERT INTO hosts(host,last_seen)
            VALUES(?,?)
            ON CONFLICT(host) DO UPDATE SET
                last_seen=excluded.last_seen
            """,
            (host, now),
        )

    def save_snapshot(
        self,
        host: str,
        payload: Any,
        *,
        kind: str = "health",
        correlation_id: str | None = None,
    ) -> int:
        now = datetime.now(timezone.utc).isoformat()
        encoded = self._encode(payload)

        with self._connect() as connection:
            self._touch_host(connection, host, now)
            cursor = connection.execute(
                """
                INSERT INTO snapshots(
                    host,created_at,kind,payload,correlation_id
                ) VALUES(?,?,?,?,?)
                """,
                (
                    host,
                    now,
                    kind,
                    encoded,
                    correlation_id,
                ),
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
            "SELECT id,host,created_at,kind,payload,correlation_id "
            "FROM snapshots WHERE lower(host)=lower(?)"
        )
        args: list[Any] = [host]
        if kind:
            sql += " AND kind=?"
            args.append(kind)
        sql += " ORDER BY id DESC LIMIT ?"
        args.append(max(1, int(limit)))

        with self._connect() as connection:
            rows = connection.execute(sql, args).fetchall()

        return [
            {
                **dict(row),
                "payload": json.loads(row["payload"]),
            }
            for row in rows
        ]

    def diff_latest(
        self,
        host: str,
        *,
        kind: str = "health",
    ) -> list[dict[str, Any]]:
        rows = self.recent_snapshots(
            host,
            kind=kind,
            limit=2,
        )
        if len(rows) < 2:
            return []
        return diff_values(
            rows[1]["payload"],
            rows[0]["payload"],
        )

    def save_job(self, record: Any) -> None:
        payload = self._jsonable(record)
        job_id = str(payload["job_id"])
        host = str(payload["host"])
        created_at = str(payload["created_at"])
        state = str(payload["state"])
        label = str(payload["label"])
        correlation_id = payload.get("correlation_id")

        with self._connect() as connection:
            self._touch_host(connection, host, created_at)
            connection.execute(
                """
                INSERT INTO jobs(
                    job_id,host,created_at,state,label,payload,correlation_id
                ) VALUES(?,?,?,?,?,?,?)
                ON CONFLICT(job_id) DO UPDATE SET
                    state=excluded.state,
                    label=excluded.label,
                    payload=excluded.payload,
                    correlation_id=excluded.correlation_id
                """,
                (
                    job_id,
                    host,
                    created_at,
                    state,
                    label,
                    self._encode(redact(payload)),
                    correlation_id,
                ),
            )

    def recent_jobs(
        self,
        host: str | None = None,
        *,
        limit: int = 50,
    ) -> list[dict[str, Any]]:
        if host:
            sql = (
                "SELECT job_id,host,created_at,state,label,payload,correlation_id "
                "FROM jobs WHERE lower(host)=lower(?) "
                "ORDER BY created_at DESC LIMIT ?"
            )
            args: tuple[Any, ...] = (
                host,
                max(1, int(limit)),
            )
        else:
            sql = (
                "SELECT job_id,host,created_at,state,label,payload,correlation_id "
                "FROM jobs ORDER BY created_at DESC LIMIT ?"
            )
            args = (max(1, int(limit)),)

        with self._connect() as connection:
            rows = connection.execute(sql, args).fetchall()

        return [
            {
                **dict(row),
                "payload": json.loads(row["payload"] or "{}"),
            }
            for row in rows
        ]

    def save_finding(
        self,
        host: str,
        finding_id: str,
        severity: str,
        payload: Any,
        *,
        correlation_id: str | None = None,
    ) -> None:
        now = datetime.now(timezone.utc).isoformat()
        with self._connect() as connection:
            self._touch_host(connection, host, now)
            connection.execute(
                """
                INSERT INTO findings(
                    host,created_at,finding_id,severity,payload,correlation_id
                ) VALUES(?,?,?,?,?,?)
                """,
                (
                    host,
                    now,
                    finding_id,
                    severity,
                    self._encode(redact(payload)),
                    correlation_id,
                ),
            )

    def save_remediation(
        self,
        host: str,
        action: str,
        success: bool,
        payload: Any,
        *,
        correlation_id: str | None = None,
    ) -> None:
        now = datetime.now(timezone.utc).isoformat()
        with self._connect() as connection:
            self._touch_host(connection, host, now)
            connection.execute(
                """
                INSERT INTO remediations(
                    host,created_at,action,success,payload,correlation_id
                ) VALUES(?,?,?,?,?,?)
                """,
                (
                    host,
                    now,
                    action,
                    int(success),
                    self._encode(payload),
                    correlation_id,
                ),
            )

    def recent_remediations(
        self,
        host: str,
        *,
        limit: int = 20,
    ) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT
                    id,host,created_at,action,success,payload,correlation_id
                FROM remediations
                WHERE lower(host)=lower(?)
                ORDER BY id DESC
                LIMIT ?
                """,
                (host, max(1, int(limit))),
            ).fetchall()

        return [
            {
                **dict(row),
                "payload": json.loads(row["payload"]),
            }
            for row in rows
        ]

    def save_execution(
        self,
        host: str,
        action: str,
        validation_state: str,
        payload: Any,
        *,
        correlation_id: str | None = None,
    ) -> None:
        now = datetime.now(timezone.utc).isoformat()
        with self._connect() as connection:
            self._touch_host(connection, host, now)
            connection.execute(
                """
                INSERT INTO executions(
                    host,created_at,action,validation_state,payload,correlation_id
                ) VALUES(?,?,?,?,?,?)
                """,
                (
                    host,
                    now,
                    action,
                    validation_state,
                    self._encode(redact(payload)),
                    correlation_id,
                ),
            )

    def save_execution_record(
        self,
        record: Any,
        *,
        correlation_id: str | None = None,
    ) -> int:
        command = record.remediation.command_result
        validation = record.remediation.validation
        now = record.finished_at or datetime.now(timezone.utc).isoformat()

        with self._connect() as connection:
            self._touch_host(connection, record.host, now)
            cursor = connection.execute(
                """
                INSERT INTO executions(
                    host,created_at,action,validation_state,payload,correlation_id,
                    operator,action_version,transport,started_at,finished_at,
                    duration_ms,risk,parameters,rollback_available,rollback_of,
                    is_rollback
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    record.host,
                    now,
                    record.action.key,
                    validation.status.value,
                    self._encode(record.audit_payload()),
                    correlation_id,
                    record.operator,
                    record.action.action_version,
                    command.transport,
                    record.started_at,
                    record.finished_at,
                    record.duration_ms,
                    record.action.risk.value,
                    self._encode(redact(record.public_parameters)),
                    int(record.rollback_available),
                    None,
                    0,
                ),
            )
            execution_id = int(cursor.lastrowid)

        record.execution_id = execution_id
        return execution_id

    def save_rollback_record(
        self,
        rollback: Any,
        *,
        original_execution_id: int,
        operator: str | None = None,
        correlation_id: str | None = None,
    ) -> int:
        now = rollback.finished_at or datetime.now(timezone.utc).isoformat()
        validation = rollback.validation
        result = rollback.result
        payload = {
            "action_key": rollback.action_key,
            "result": {
                "success": result.success,
                "transport": result.transport,
                "return_code": result.return_code,
                "duration_ms": result.duration_ms,
                "indeterminate": result.indeterminate,
                "error": result.stderr,
                "data": result.data,
            },
            "validation": {
                "status": validation.status.value,
                "message": validation.message,
                "evidence": validation.evidence,
            },
            "started_at": rollback.started_at,
            "finished_at": rollback.finished_at,
            "duration_ms": rollback.duration_ms,
        }

        with self._connect() as connection:
            self._touch_host(connection, rollback.host, now)
            cursor = connection.execute(
                """
                INSERT INTO executions(
                    host,created_at,action,validation_state,payload,correlation_id,
                    operator,action_version,transport,started_at,finished_at,
                    duration_ms,risk,parameters,rollback_available,rollback_of,
                    is_rollback
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    rollback.host,
                    now,
                    f"rollback:{rollback.action_key}",
                    validation.status.value,
                    self._encode(redact(payload)),
                    correlation_id,
                    operator,
                    1,
                    result.transport,
                    rollback.started_at,
                    rollback.finished_at,
                    rollback.duration_ms,
                    "ROLLBACK",
                    self._encode({}),
                    0,
                    original_execution_id,
                    1,
                ),
            )
            if validation.status.value == "PASS":
                connection.execute(
                    """
                    UPDATE executions
                    SET rollback_available=0
                    WHERE id=?
                    """,
                    (original_execution_id,),
                )
            return int(cursor.lastrowid)

    def recent_executions(
        self,
        host: str,
        *,
        limit: int = 30,
    ) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT
                    id,host,created_at,action,validation_state,payload,correlation_id,
                    operator,action_version,transport,started_at,finished_at,
                    duration_ms,risk,parameters,rollback_available,rollback_of,
                    is_rollback
                FROM executions
                WHERE lower(host)=lower(?)
                ORDER BY id DESC
                LIMIT ?
                """,
                (host, max(1, int(limit))),
            ).fetchall()

        return [
            {
                **dict(row),
                "payload": json.loads(row["payload"]),
            }
            for row in rows
        ]

    def save_report(
        self,
        host: str,
        fmt: str,
        payload: str,
        *,
        path: str | None = None,
        correlation_id: str | None = None,
    ) -> None:
        now = datetime.now(timezone.utc).isoformat()
        with self._connect() as connection:
            self._touch_host(connection, host, now)
            connection.execute(
                """
                INSERT INTO reports(
                    host,created_at,format,path,payload,correlation_id
                ) VALUES(?,?,?,?,?,?)
                """,
                (
                    host,
                    now,
                    fmt,
                    path,
                    payload,
                    correlation_id,
                ),
            )

    def prune(self, retention_days: int) -> dict[str, int]:
        cutoff = (
            datetime.now(timezone.utc)
            - timedelta(days=max(1, int(retention_days)))
        ).isoformat()

        deleted: dict[str, int] = {}
        with self._connect() as connection:
            for table in (
                "snapshots",
                "jobs",
                "findings",
                "remediations",
                "executions",
                "reports",
            ):
                cursor = connection.execute(
                    f"DELETE FROM {table} WHERE created_at < ?",
                    (cutoff,),
                )
                deleted[table] = max(0, cursor.rowcount)

            connection.execute(
                """
                DELETE FROM hosts
                WHERE last_seen < ?
                  AND host NOT IN (SELECT host FROM snapshots)
                  AND host NOT IN (SELECT host FROM jobs)
                  AND host NOT IN (SELECT host FROM findings)
                  AND host NOT IN (SELECT host FROM remediations)
                  AND host NOT IN (SELECT host FROM executions)
                  AND host NOT IN (SELECT host FROM reports)
                """,
                (cutoff,),
            )

        return deleted
