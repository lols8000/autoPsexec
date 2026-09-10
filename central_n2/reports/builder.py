from __future__ import annotations

from datetime import datetime
from typing import Any


class SupportReportBuilder:
    def build(
        self,
        *,
        host: str,
        user: str | None = None,
        problem: str | None = None,
        diagnosis: str | None = None,
        actions: list[str] | None = None,
        validation: Any = None,
        result: str | None = None,
    ) -> dict[str, Any]:
        return {
            "generated_at": datetime.now().astimezone().isoformat(
                timespec="seconds"
            ),
            "host": host,
            "user": user,
            "problem": problem,
            "diagnosis": diagnosis,
            "actions": actions or [],
            "validation": validation,
            "result": result,
        }
