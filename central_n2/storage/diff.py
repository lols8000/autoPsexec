from __future__ import annotations

from typing import Any


def diff_values(
    before: Any,
    after: Any,
    path: str = "",
) -> list[dict[str, Any]]:
    changes: list[dict[str, Any]] = []

    if isinstance(before, dict) and isinstance(after, dict):
        for key in sorted(set(before) | set(after)):
            child = f"{path}.{key}" if path else str(key)
            if key not in before:
                changes.append(
                    {
                        "path": child,
                        "type": "added",
                        "before": None,
                        "after": after[key],
                    }
                )
            elif key not in after:
                changes.append(
                    {
                        "path": child,
                        "type": "removed",
                        "before": before[key],
                        "after": None,
                    }
                )
            else:
                changes.extend(
                    diff_values(
                        before[key],
                        after[key],
                        child,
                    )
                )
        return changes

    if before != after:
        changes.append(
            {
                "path": path or "$",
                "type": "changed",
                "before": before,
                "after": after,
            }
        )
    return changes
