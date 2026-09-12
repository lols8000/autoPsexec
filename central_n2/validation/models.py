from __future__ import annotations

import hashlib
from dataclasses import asdict, dataclass, field
from enum import Enum
from typing import Any


class CheckState(str, Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    UNKNOWN = "UNKNOWN"
    SKIP = "SKIP"


@dataclass(frozen=True, slots=True)
class ControlledActionSpec:
    key: str
    parameters: dict[str, Any] = field(default_factory=dict)
    rollback_after: bool = False


@dataclass(frozen=True, slots=True)
class EndpointSpec:
    alias: str
    target: str
    enabled: bool = True
    expected_state: str | None = None
    expected_transport: str | None = None
    required_capabilities: tuple[str, ...] = ()
    roles: tuple[str, ...] = ()
    actions: tuple[ControlledActionSpec, ...] = ()


@dataclass(frozen=True, slots=True)
class ValidationCheck:
    key: str
    state: CheckState
    message: str
    evidence: Any = None
    required: bool = True


@dataclass(slots=True)
class EndpointValidationResult:
    alias: str
    target_fingerprint: str
    correlation_id: str
    started_at: str
    finished_at: str
    checks: list[ValidationCheck] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def status(self) -> CheckState:
        required = [item for item in self.checks if item.required]
        if any(item.state is CheckState.FAIL for item in required):
            return CheckState.FAIL
        if any(item.state is CheckState.UNKNOWN for item in required):
            return CheckState.UNKNOWN
        if required and all(
            item.state in {CheckState.PASS, CheckState.SKIP}
            for item in required
        ):
            return CheckState.PASS
        return CheckState.UNKNOWN

    def public_dict(self) -> dict[str, Any]:
        return {
            "alias": self.alias,
            "target_fingerprint": self.target_fingerprint,
            "correlation_id": self.correlation_id,
            "started_at": self.started_at,
            "finished_at": self.finished_at,
            "status": self.status.value,
            "metadata": self.metadata,
            "checks": [
                {
                    **asdict(item),
                    "state": item.state.value,
                }
                for item in self.checks
            ],
        }


@dataclass(slots=True)
class CampaignResult:
    name: str
    started_at: str
    finished_at: str
    endpoints: list[EndpointValidationResult]

    @property
    def status(self) -> CheckState:
        states = [item.status for item in self.endpoints]
        if any(state is CheckState.FAIL for state in states):
            return CheckState.FAIL
        if any(state is CheckState.UNKNOWN for state in states):
            return CheckState.UNKNOWN
        if states and all(state is CheckState.PASS for state in states):
            return CheckState.PASS
        return CheckState.UNKNOWN

    def public_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "started_at": self.started_at,
            "finished_at": self.finished_at,
            "status": self.status.value,
            "endpoints": [
                endpoint.public_dict()
                for endpoint in self.endpoints
            ],
        }


def fingerprint_target(target: str) -> str:
    return hashlib.sha256(
        target.strip().casefold().encode("utf-8")
    ).hexdigest()[:12]
