from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Callable

from .models import ExecutionAction, PrivilegeLevel


class PolicyState(str, Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    WARN = "WARN"


@dataclass(frozen=True, slots=True)
class PolicyCheck:
    key: str
    state: PolicyState
    message: str
    evidence: Any = None


@dataclass(slots=True)
class ExecutionPolicyContext:
    host: str
    ready: bool
    transport: str
    capabilities: dict[str, Any] = field(default_factory=dict)
    capability_error: str | None = None

    @classmethod
    def from_session(cls, session: Any) -> "ExecutionPolicyContext":
        return cls(
            host=str(session.host),
            ready=bool(session.ready),
            transport=str(session.transport or "unknown").casefold(),
            capabilities=dict(session.capabilities or {}),
            capability_error=session.capability_error,
        )


@dataclass(slots=True)
class ExecutionPolicyReport:
    allowed: bool
    checks: list[PolicyCheck] = field(default_factory=list)

    @property
    def failures(self) -> list[PolicyCheck]:
        return [
            item
            for item in self.checks
            if item.state is PolicyState.FAIL
        ]


CustomPrecondition = Callable[
    [ExecutionPolicyContext, dict[str, Any]],
    list[PolicyCheck],
]


class ExecutionPolicy:
    def evaluate(
        self,
        action: ExecutionAction,
        context: ExecutionPolicyContext,
        parameters: dict[str, Any],
        *,
        custom_preconditions: tuple[CustomPrecondition, ...] = (),
    ) -> ExecutionPolicyReport:
        checks: list[PolicyCheck] = []

        checks.append(
            PolicyCheck(
                "session.ready",
                PolicyState.PASS if context.ready else PolicyState.FAIL,
                (
                    "Transporte administrativo validado."
                    if context.ready
                    else "Nenhum transporte administrativo validado."
                ),
                context.transport,
            )
        )

        transport = context.transport.casefold()
        allowed = {
            item.casefold()
            for item in action.allowed_transports
        }
        checks.append(
            PolicyCheck(
                "transport",
                PolicyState.PASS if transport in allowed else PolicyState.FAIL,
                (
                    f"Transporte {transport} permitido."
                    if transport in allowed
                    else (
                        f"Transporte {transport} não é permitido; "
                        f"aceitos: {', '.join(action.allowed_transports)}."
                    )
                ),
                transport,
            )
        )

        for capability in action.required_capabilities:
            value = context.capabilities.get(capability)
            checks.append(
                PolicyCheck(
                    f"capability.{capability}",
                    PolicyState.PASS if bool(value) else PolicyState.FAIL,
                    (
                        f"Capability {capability} disponível."
                        if bool(value)
                        else f"Capability {capability} indisponível."
                    ),
                    value,
                )
            )

        privilege = action.required_privilege
        is_admin = context.capabilities.get("IsAdmin")
        is_system = context.capabilities.get("IsSystem")

        if privilege is PrivilegeLevel.ANY:
            state = PolicyState.PASS
            message = "A ação não exige contexto privilegiado específico."
        elif privilege is PrivilegeLevel.ADMIN:
            state = (
                PolicyState.PASS
                if is_admin is True
                else PolicyState.FAIL
            )
            message = (
                "Contexto administrativo confirmado."
                if state is PolicyState.PASS
                else "Contexto administrativo não foi confirmado."
            )
        elif privilege is PrivilegeLevel.SYSTEM:
            state = (
                PolicyState.PASS
                if is_system is True
                else PolicyState.FAIL
            )
            message = (
                "Contexto SYSTEM confirmado."
                if state is PolicyState.PASS
                else "A ação exige contexto SYSTEM."
            )
        else:
            state = (
                PolicyState.PASS
                if is_admin is True and is_system is False
                else PolicyState.FAIL
            )
            message = (
                "Contexto administrativo de usuário confirmado."
                if state is PolicyState.PASS
                else (
                    "A ação exige contexto de usuário administrativo; "
                    "SYSTEM não é aceito."
                )
            )

        checks.append(
            PolicyCheck(
                "privilege",
                state,
                message,
                {
                    "required": privilege.value,
                    "is_admin": is_admin,
                    "is_system": is_system,
                    "identity": context.capabilities.get("ExecutionIdentity"),
                },
            )
        )

        if context.capability_error:
            checks.append(
                PolicyCheck(
                    "capabilities.probe",
                    PolicyState.WARN,
                    "O probe de capabilities retornou erro parcial.",
                    context.capability_error,
                )
            )

        for precondition in custom_preconditions:
            checks.extend(precondition(context, parameters))

        return ExecutionPolicyReport(
            allowed=not any(
                item.state is PolicyState.FAIL
                for item in checks
            ),
            checks=checks,
        )
