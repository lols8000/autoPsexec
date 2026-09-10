from __future__ import annotations

from dataclasses import asdict
from typing import Any

from core.evaluation import EvaluationPolicy, evaluate_snapshot, score_checks


def evaluate_compliance(
    snapshot: dict[str, Any],
    baseline: dict[str, Any],
) -> dict[str, Any]:
    policy = EvaluationPolicy.from_mapping(baseline)
    checks = evaluate_snapshot(snapshot, policy)
    summary = score_checks(checks)

    items: list[dict[str, Any]] = []
    for check in checks:
        item = asdict(check)
        item["state"] = check.state.value
        item["compliant"] = check.compliant
        items.append(item)

    return {
        "score": summary["compliance_score"],
        "compliant": summary["passed"],
        "failed": summary["failed"],
        "unknown": summary["unknown"],
        "total": len(checks),
        "overall_state": summary["overall_state"],
        "items": items,
    }
