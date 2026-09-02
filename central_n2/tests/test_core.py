from __future__ import annotations

import json
from pathlib import Path

from core.logger import AuditLogger
from core.result import CommandResult
from modules.batch import BatchRunner


def test_command_result_failure():
    result = CommandResult.failure("PC01", "teste", "erro")
    assert result.success is False
    assert result.host == "PC01"
    assert result.stderr == "erro"
    assert result.return_code == 1


def test_audit_logger_writes_jsonl(tmp_path: Path):
    logger = AuditLogger(tmp_path)
    result = CommandResult(True, "hostname", "PC01", stdout="PC01")
    logger.log_result("hostname", result)

    files = list(tmp_path.glob("*.jsonl"))
    assert len(files) == 1
    payload = json.loads(files[0].read_text(encoding="utf-8").strip())
    assert payload["action"] == "hostname"
    assert payload["host"] == "PC01"
    assert payload["success"] is True


def test_batch_runner_collects_success_and_failure():
    runner = BatchRunner(max_workers=2)

    def action(host: str) -> CommandResult:
        if host == "PC02":
            return CommandResult.failure(host, "ping", "offline")
        return CommandResult(True, "ping", host, stdout="ok")

    results = runner.run(["PC02", "PC01"], action)
    assert [r["host"] for r in results] == ["PC01", "PC02"]
    assert results[0]["success"] is True
    assert results[1]["success"] is False
