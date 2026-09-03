from __future__ import annotations
from dataclasses import dataclass, field
from enum import Enum
from typing import Any

class Severity(str, Enum):
    INFO="info"; LOW="low"; MEDIUM="medium"; HIGH="high"; CRITICAL="critical"

@dataclass(slots=True)
class Finding:
    id: str
    severity: Severity
    message: str
    evidence: dict[str, Any] = field(default_factory=dict)
    recommendations: list[str] = field(default_factory=list)

@dataclass(slots=True)
class Diagnosis:
    code: str
    title: str
    confidence: str
    rationale: str
    findings: list[Finding] = field(default_factory=list)
    actions: list[str] = field(default_factory=list)
