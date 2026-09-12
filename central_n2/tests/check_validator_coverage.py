from __future__ import annotations

import json
import sys
from pathlib import Path


MINIMUM = 80.0
TARGETS = (
    "execution/validators.py",
    "remediation/validators.py",
)


def main() -> int:
    if len(sys.argv) != 2:
        print("usage: check_validator_coverage.py <coverage.json>")
        return 2

    report_path = Path(sys.argv[1])
    payload = json.loads(report_path.read_text(encoding="utf-8"))
    files = payload.get("files", {})
    failures: list[str] = []

    for target in TARGETS:
        data = files.get(target)
        if data is None:
            failures.append(f"{target}: ausente no relatório de coverage")
            continue

        summary = data.get("summary", {})
        percent = float(summary.get("percent_covered", 0.0))
        print(f"{target}: {percent:.2f}% (mínimo {MINIMUM:.2f}%)")
        if percent < MINIMUM:
            failures.append(
                f"{target}: {percent:.2f}% < {MINIMUM:.2f}%"
            )

    if failures:
        print("\nValidator coverage gate FAILED:")
        for failure in failures:
            print(f" - {failure}")
        return 1

    print("\nValidator coverage gate PASSED.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
