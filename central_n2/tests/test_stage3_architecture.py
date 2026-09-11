from __future__ import annotations

import json
import sqlite3
from pathlib import Path

from core.actions import ActionSpec
from core.jobs import JobManager, OperationClass
from core.logger import AuditLogger
from core.result import CommandResult
from remediation import (
    RemediationEngine,
    RemediationSpec,
    ValidationStatus,
    validate_spooler,
)
from storage.database import CentralDatabase
from ui.console_base import ConsoleBase
from ui.console_v5 import ConsoleUIV5


def test_v5_uses_clean_console_base_not_v3():
    assert issubclass(ConsoleUIV5, ConsoleBase)
    assert ConsoleUIV5.__mro__[1] is ConsoleBase


def test_action_spec_is_independent_from_display_title():
    spec = ActionSpec(
        key="system.restart",
        title="Qualquer texto traduzido",
        operation_class=OperationClass.DISRUPTIVE,
        requires_confirmation=True,
    )
    assert spec.key == "system.restart"
    assert spec.operation_class is OperationClass.DISRUPTIVE
    assert spec.requires_confirmation is True


def test_responsive_runner_can_share_single_job_manager():
    manager = JobManager(max_workers=1, heartbeat_seconds=0.01)
    from core.jobs import ResponsiveJobRunner

    runner = ResponsiveJobRunner(manager=manager)
    try:
        assert runner.run(
            "consulta",
            lambda: "ok",
            host="PC01",
        ) == "ok"
        assert len(manager.list_records()) == 1
    finally:
        runner.shutdown()
        manager.shutdown()


def test_database_uses_versioned_schema_and_correlation(tmp_path: Path):
    database = CentralDatabase(tmp_path / "central.db")

    with sqlite3.connect(database.path) as connection:
        version = connection.execute(
            "PRAGMA user_version"
        ).fetchone()[0]
        columns = {
            row[1]
            for row in connection.execute(
                "PRAGMA table_info(snapshots)"
            ).fetchall()
        }

    assert version == CentralDatabase.SCHEMA_VERSION
    assert "correlation_id" in columns
    with sqlite3.connect(database.path) as connection:
        execution_tables = {
            row[0]
            for row in connection.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
    assert "executions" in execution_tables
    assert database.integrity_check() is True

    database.save_snapshot(
        "PC01",
        {"ok": True},
        correlation_id="ATD001",
    )
    row = database.recent_snapshots("PC01", limit=1)[0]
    assert row["correlation_id"] == "ATD001"


def test_database_prunes_all_operational_tables(tmp_path: Path):
    database = CentralDatabase(tmp_path / "central.db")
    old = "2000-01-01T00:00:00+00:00"

    with database._connect() as connection:
        connection.execute(
            "INSERT INTO hosts(host,last_seen) VALUES(?,?)",
            ("OLDPC", old),
        )
        connection.execute(
            "INSERT INTO snapshots(host,created_at,kind,payload) "
            "VALUES(?,?,?,?)",
            ("OLDPC", old, "health", "{}"),
        )
        connection.execute(
            "INSERT INTO jobs(job_id,host,created_at,state,label,payload) "
            "VALUES(?,?,?,?,?,?)",
            ("oldjob", "OLDPC", old, "SUCCESS", "x", "{}"),
        )
        connection.execute(
            "INSERT INTO findings(host,created_at,finding_id,severity,payload) "
            "VALUES(?,?,?,?,?)",
            ("OLDPC", old, "x", "low", "{}"),
        )
        connection.execute(
            "INSERT INTO remediations(host,created_at,action,success,payload) "
            "VALUES(?,?,?,?,?)",
            ("OLDPC", old, "x", 1, "{}"),
        )
        connection.execute(
            "INSERT INTO executions(host,created_at,action,validation_state,payload) "
            "VALUES(?,?,?,?,?)",
            ("OLDPC", old, "x", "PASS", "{}"),
        )
        connection.execute(
            "INSERT INTO reports(host,created_at,format,payload) "
            "VALUES(?,?,?,?)",
            ("OLDPC", old, "json", "{}"),
        )

    deleted = database.prune(1)
    assert all(
        deleted[name] == 1
        for name in (
            "snapshots",
            "jobs",
            "findings",
            "remediations",
            "executions",
            "reports",
        )
    )


def test_audit_logger_is_compact_by_default(tmp_path: Path):
    logger = AuditLogger(tmp_path)
    result = CommandResult(
        True,
        "powershell.exe -EncodedCommand SECRET",
        "PC01",
        stdout="large diagnostic payload",
        data={"token": "secret"},
        transport="winrm",
    )
    logger.log_result("health.collect", result)

    line = next(tmp_path.glob("*.jsonl")).read_text(
        encoding="utf-8"
    ).splitlines()[0]
    payload = json.loads(line)

    assert "command" not in payload
    assert "stdout" not in payload
    assert "data" not in payload
    assert payload["transport"] == "winrm"


def test_remediation_uses_specific_validator():
    states = iter(
        [
            {"Status": "Stopped"},
            {"Status": "Running"},
        ]
    )

    result = RemediationEngine().execute(
        "PC01",
        RemediationSpec(
            "restart_spooler",
            "Reiniciar Spooler",
            "baixo",
        ),
        lambda host: CommandResult(
            True,
            "restart",
            host,
            data={"Status": "Running"},
        ),
        before_probe=lambda host: next(states),
        after_probe=lambda host: next(states),
        validator=validate_spooler,
    )

    assert result.validation.status is ValidationStatus.PASS



class _BootstrapExecutor:
    logger = None


def test_v5_bootstrap_accepts_injected_settings_and_has_single_state(tmp_path: Path):
    config_dir = tmp_path / "config"
    config_dir.mkdir()
    settings_path = config_dir / "settings.json"
    settings_path.write_text("{}", encoding="utf-8")

    settings = {
        "runtime": {
            "max_workers": 1,
        },
        "ui": {
            "heartbeat_seconds": 0.01,
            "long_operation_timeout_seconds": 30,
        },
        "persistence": {
            "enabled": False,
        },
        "compliance": {
            "profile": "DEFAULT",
        },
        "updates": {
            "enabled": False,
            "repository": "lols8000/autoPsexec",
        },
    }

    ui = ConsoleUIV5(
        _BootstrapExecutor(),
        settings_path,
        settings=settings,
    )
    try:
        assert ui.settings is settings
        assert ui.context.host is None
        assert ui.context.session is None
        assert ui.context.health_snapshot is None
        assert ui.context.diagnoses == []
        assert ui.context.playbook is None
        assert ui.context.remediation is None
        assert ui.context.execution is None
        assert ui.context.rollback_stack == []
        assert ui.context.report_path is None
        assert ui.updates_enabled is False

        legacy_mirrors = {
            "current_session",
            "health_snapshot",
            "last_diagnoses",
            "last_playbook",
            "last_remediation",
            "last_report_path",
            "correlation_id",
        }
        assert legacy_mirrors.isdisjoint(vars(ui))
    finally:
        ui.jobs.shutdown()
        ui.job_manager.shutdown()



def test_database_persists_execution_history(tmp_path: Path):
    database = CentralDatabase(tmp_path / "central.db")
    database.save_execution(
        "PC01",
        "network.flush_dns",
        "PASS",
        {"action": "network.flush_dns", "result": "ok"},
        correlation_id="EXEC001",
    )

    rows = database.recent_executions("PC01", limit=5)

    assert len(rows) == 1
    assert rows[0]["action"] == "network.flush_dns"
    assert rows[0]["validation_state"] == "PASS"
    assert rows[0]["correlation_id"] == "EXEC001"
    assert rows[0]["payload"]["result"] == "ok"



def test_attendance_context_clears_rollback_stack_when_host_changes():
    from core.context import AttendanceContext

    context = AttendanceContext.start("PC01", session=object())
    context.rollback_stack.append(object())
    context.execution = object()

    context.bind("PC02", session=object())

    assert context.host == "PC02"
    assert context.execution is None
    assert context.rollback_stack == []
