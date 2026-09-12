from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class CaseStatus(str, Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    UNKNOWN = "UNKNOWN"
    SKIP = "SKIP"


@dataclass(frozen=True, slots=True)
class QualificationCase:
    key: str
    title: str
    category: str
    kind: str
    required: bool = True
    capability: str | None = None
    domain_member: bool = False


@dataclass(slots=True)
class QualificationCaseResult:
    key: str
    title: str
    category: str
    status: CaseStatus
    started_at: str
    finished_at: str
    duration_ms: int
    transport: str | None = None
    message: str = ""
    evidence: Any = None


@dataclass(slots=True)
class QualificationReport:
    host: str
    profile: str
    started_at: str
    finished_at: str
    transport: str
    connectivity_state: str
    cases: list[QualificationCaseResult] = field(default_factory=list)
    requested_actions: list[str] = field(default_factory=list)
    operator: str | None = None
    version: str | None = None

    def _case(self, key: str) -> QualificationCaseResult | None:
        return next(
            (case for case in self.cases if case.key == key),
            None,
        )

    def _profile_requirement_met(self) -> bool:
        specialized_cases = {
            "printer": "printers.inventory",
            "glpi": "glpi.status",
            "domain": "domain.status",
        }
        case_key = specialized_cases.get(self.profile)
        if case_key:
            case = self._case(case_key)
            return case is not None and case.status is CaseStatus.PASS

        if self.profile == "notebook":
            capabilities = self._case("core.capabilities")
            evidence = capabilities.evidence if capabilities else None
            return (
                isinstance(evidence, dict)
                and evidence.get("Battery") is True
            )

        return True

    @property
    def passed(self) -> bool:
        no_unresolved_failure = not any(
            case.status in {CaseStatus.FAIL, CaseStatus.UNKNOWN}
            for case in self.cases
        )
        return no_unresolved_failure and self._profile_requirement_met()

    def summary(self) -> dict[str, int]:
        counts = {status.value: 0 for status in CaseStatus}
        for case in self.cases:
            counts[case.status.value] += 1
        return counts
