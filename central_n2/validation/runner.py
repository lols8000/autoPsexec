from __future__ import annotations

import json
import time
import uuid
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

from core.result import CommandResult
from core.session import SessionManager, WorkstationSession
from execution import (
    ExecutionBlockedError,
    ExecutionDependencies,
    ExecutionEngine,
    ExecutionPolicyContext,
    RecoveryResult,
    RiskLevel,
    build_execution_registry,
)
from modules.certificates import CertificatesModule
from modules.devices import DevicesModule
from modules.disk import DiskModule
from modules.domain import DomainModule
from modules.file_ops import FileOperationsModule
from modules.glpi import GLPIModule
from modules.health import HealthModule
from modules.network import NetworkModule
from modules.packages import PackagesModule
from modules.printers import PrintersModule
from modules.registry_actions import RegistryActionsModule
from modules.repair import RepairModule
from modules.security import SecurityModule
from modules.software import SoftwareModule
from modules.system import SystemModule
from modules.updates import UpdatesModule
from modules.users_profiles import UsersProfilesModule

from .models import (
    CampaignResult,
    CheckState,
    ControlledActionSpec,
    EndpointSpec,
    EndpointValidationResult,
    ValidationCheck,
    fingerprint_target,
)


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _redact_target(value: Any, target: str) -> Any:
    if isinstance(value, str):
        if not target:
            return value
        return value.replace(target, "<TARGET>").replace(
            target.casefold(),
            "<TARGET>",
        )
    if isinstance(value, dict):
        return {
            str(key): _redact_target(item, target)
            for key, item in value.items()
        }
    if isinstance(value, list):
        return [_redact_target(item, target) for item in value]
    if isinstance(value, tuple):
        return tuple(_redact_target(item, target) for item in value)
    return value


def _sanitize_checks(
    checks: list[ValidationCheck],
    target: str,
) -> list[ValidationCheck]:
    return [
        ValidationCheck(
            key=item.key,
            state=item.state,
            message=_redact_target(item.message, target),
            evidence=_redact_target(item.evidence, target),
            required=item.required,
        )
        for item in checks
    ]


def _check_from_result(
    key: str,
    result: CommandResult,
    success_message: str,
    *,
    required: bool = True,
) -> ValidationCheck:
    if result.success:
        return ValidationCheck(
            key=key,
            state=CheckState.PASS,
            message=success_message,
            evidence={"transport": result.transport},
            required=required,
        )
    return ValidationCheck(
        key=key,
        state=(
            CheckState.UNKNOWN
            if result.indeterminate
            else CheckState.FAIL
        ),
        message=result.stderr or "Consulta falhou.",
        evidence={
            "transport": result.transport,
            "indeterminate": result.indeterminate,
        },
        required=required,
    )


class EndpointValidationRunner:
    def __init__(
        self,
        executor,
        settings: dict[str, Any],
        settings_path: Path,
        *,
        sleep: Callable[[float], None] = time.sleep,
    ) -> None:
        self.executor = executor
        self.settings = settings
        self.settings_path = settings_path
        self.sleep = sleep
        self.sessions = SessionManager(executor)

        self.health = HealthModule(executor)
        self.network = NetworkModule(executor)
        self.printers = PrintersModule(executor)
        self.domain = DomainModule(executor)
        self.glpi = GLPIModule(
            executor,
            settings_path,
            settings=settings,
        )
        self.system = SystemModule(executor)
        self.software = SoftwareModule(
            executor,
            settings_path,
            settings=settings,
        )
        self.devices = DevicesModule(executor)
        self.users = UsersProfilesModule(executor)
        self.disk = DiskModule(executor)
        self.security = SecurityModule(executor)
        self.updates = UpdatesModule(executor)
        self.repair = RepairModule(executor)
        self.packages = PackagesModule(executor, settings)
        self.certificates = CertificatesModule(executor, settings)
        self.registry_actions = RegistryActionsModule(executor, settings)
        execution_settings = settings.get("execution", {})
        self.files = FileOperationsModule(
            executor,
            allowed_roots=execution_settings.get("file_roots"),
        )

        deps = ExecutionDependencies(
            system=self.system,
            network=self.network,
            software=self.software,
            printers=self.printers,
            devices=self.devices,
            domain=self.domain,
            users=self.users,
            disk=self.disk,
            glpi=self.glpi,
            health=self.health,
            security=self.security,
            updates=self.updates,
            repair=self.repair,
            packages=self.packages,
            certificates=self.certificates,
            registry_actions=self.registry_actions,
            files=self.files,
        )
        self.registry = build_execution_registry(deps)
        self.engine = ExecutionEngine(self.registry)

    @staticmethod
    def _metadata(session: WorkstationSession) -> dict[str, Any]:
        allowed = (
            "OS",
            "Build",
            "Architecture",
            "Manufacturer",
            "Model",
            "DomainMember",
            "IsAdmin",
            "IsSystem",
            "Battery",
            "GLPI",
        )
        return {
            key: session.capabilities.get(key)
            for key in allowed
            if key in session.capabilities
        }

    def _safe_checks(
        self,
        spec: EndpointSpec,
        session: WorkstationSession,
    ) -> list[ValidationCheck]:
        checks: list[ValidationCheck] = []
        ready = session.ready
        checks.append(
            ValidationCheck(
                key="session.ready",
                state=CheckState.PASS if ready else CheckState.FAIL,
                message=(
                    f"Sessão pronta: {session.connectivity.get('state')}."
                    if ready
                    else (
                        "Nenhum transporte administrativo utilizável: "
                        f"{session.connectivity.get('state')}."
                    )
                ),
                evidence={
                    "state": session.connectivity.get("state"),
                    "transport": session.transport,
                    "dns": session.connectivity.get("dns"),
                    "tcp_445": session.connectivity.get("tcp_445"),
                    "tcp_5985": session.connectivity.get("tcp_5985"),
                    "tcp_5986": session.connectivity.get("tcp_5986"),
                },
            )
        )

        if spec.expected_state:
            actual = str(session.connectivity.get("state") or "")
            checks.append(
                ValidationCheck(
                    key="expectation.state",
                    state=(
                        CheckState.PASS
                        if actual == spec.expected_state
                        else CheckState.FAIL
                    ),
                    message=(
                        f"Estado esperado confirmado: {actual}."
                        if actual == spec.expected_state
                        else (
                            f"Esperado {spec.expected_state}; "
                            f"obtido {actual or '-'}."
                        )
                    ),
                    evidence={
                        "expected": spec.expected_state,
                        "actual": actual,
                    },
                )
            )

        if spec.expected_transport:
            actual_transport = session.transport.casefold()
            expected_transport = spec.expected_transport.casefold()
            checks.append(
                ValidationCheck(
                    key="expectation.transport",
                    state=(
                        CheckState.PASS
                        if actual_transport == expected_transport
                        else CheckState.FAIL
                    ),
                    message=(
                        "Transporte esperado confirmado: "
                        f"{session.transport}."
                        if actual_transport == expected_transport
                        else (
                            f"Esperado {spec.expected_transport}; "
                            f"obtido {session.transport}."
                        )
                    ),
                    evidence={
                        "expected": spec.expected_transport,
                        "actual": session.transport,
                    },
                )
            )

        for capability in spec.required_capabilities:
            value = session.capabilities.get(capability)
            checks.append(
                ValidationCheck(
                    key=f"capability.{capability}",
                    state=(
                        CheckState.PASS
                        if bool(value)
                        else CheckState.FAIL
                    ),
                    message=(
                        f"Capability {capability} disponível."
                        if bool(value)
                        else f"Capability {capability} indisponível."
                    ),
                    evidence=value,
                )
            )

        if not ready:
            return checks

        checks.extend(
            (
                _check_from_result(
                    "probe.health",
                    self.health.snapshot(spec.target),
                    "Health snapshot coletado.",
                ),
                _check_from_result(
                    "probe.network.adapters",
                    self.network.adapters(spec.target),
                    "Inventário de adaptadores coletado.",
                ),
                _check_from_result(
                    "probe.network.ip",
                    self.network.ip_configuration(spec.target),
                    "Configuração IP coletada.",
                ),
            )
        )

        roles = {role.casefold() for role in spec.roles}
        if "domain" in roles:
            checks.append(
                _check_from_result(
                    "probe.domain",
                    self.domain.status(spec.target),
                    "Estado de domínio coletado.",
                )
            )
        if "printer" in roles:
            checks.extend(
                (
                    _check_from_result(
                        "probe.printers",
                        self.printers.list(spec.target),
                        "Inventário de impressoras coletado.",
                    ),
                    _check_from_result(
                        "probe.spooler",
                        self.printers.spooler_status(spec.target),
                        "Estado do Spooler coletado.",
                    ),
                )
            )
        if "glpi" in roles:
            checks.append(
                _check_from_result(
                    "probe.glpi",
                    self.glpi.status(spec.target),
                    "Estado do GLPI Agent coletado.",
                )
            )
        if "pnp" in roles:
            checks.append(
                _check_from_result(
                    "probe.pnp",
                    self.devices.present_devices(spec.target),
                    "Inventário PnP coletado.",
                )
            )
        if "laptop" in roles:
            battery = session.capabilities.get("Battery")
            checks.append(
                ValidationCheck(
                    key="expectation.battery",
                    state=(
                        CheckState.PASS
                        if battery is True
                        else CheckState.FAIL
                    ),
                    message=(
                        "Bateria detectada."
                        if battery is True
                        else "Perfil laptop sem bateria detectável."
                    ),
                    evidence=battery,
                )
            )

        return checks

    def _recover(
        self,
        host: str,
        timeout_seconds: int,
        delay_seconds: int,
    ) -> RecoveryResult:
        if delay_seconds > 0:
            self.sleep(delay_seconds)

        started = time.monotonic()
        attempts = 0
        last_error: str | None = None
        last_state: str | None = None

        while time.monotonic() - started < timeout_seconds:
            attempts += 1
            try:
                session = self.sessions.open(host, refresh=True)
                last_state = str(
                    session.connectivity.get("state") or ""
                )
                if session.ready:
                    return RecoveryResult(
                        attempted=True,
                        ready=True,
                        attempts=attempts,
                        elapsed_seconds=round(
                            time.monotonic() - started,
                            2,
                        ),
                        transport=session.transport,
                        state=last_state,
                    )
            except Exception as exc:
                last_error = f"{type(exc).__name__}: {exc}"
            self.sleep(5)

        return RecoveryResult(
            attempted=True,
            ready=False,
            attempts=attempts,
            elapsed_seconds=round(
                time.monotonic() - started,
                2,
            ),
            state=last_state,
            error=last_error or "Prazo de recuperação excedido.",
        )

    def _controlled_checks(
        self,
        spec: EndpointSpec,
        session: WorkstationSession,
        *,
        allow_mutations: bool,
        allow_high_risk: bool,
        allow_reboot: bool,
        operator: str,
    ) -> list[ValidationCheck]:
        checks: list[ValidationCheck] = []

        for requested in spec.actions:
            check_key = f"action.{requested.key}"
            if not allow_mutations:
                checks.append(
                    ValidationCheck(
                        key=check_key,
                        state=CheckState.SKIP,
                        message=(
                            "Ação controlada não executada; "
                            "--allow-mutations ausente."
                        ),
                        required=False,
                    )
                )
                continue

            try:
                bound = self.registry.get(requested.key)
            except KeyError as exc:
                checks.append(
                    ValidationCheck(
                        key=check_key,
                        state=CheckState.FAIL,
                        message=str(exc),
                    )
                )
                continue

            action = bound.spec
            if action.risk is RiskLevel.CRITICAL:
                if requested.key != "energy.restart" or not allow_reboot:
                    checks.append(
                        ValidationCheck(
                            key=check_key,
                            state=CheckState.FAIL,
                            message=(
                                "Ação CRITICAL bloqueada pelo runner. "
                                "Somente energy.restart pode ser homologado "
                                "com --allow-reboot."
                            ),
                        )
                    )
                    continue
            elif (
                action.risk is RiskLevel.HIGH
                and not allow_high_risk
            ):
                checks.append(
                    ValidationCheck(
                        key=check_key,
                        state=CheckState.FAIL,
                        message="Ação HIGH exige --allow-high-risk.",
                    )
                )
                continue

            if (
                action.destructive
                and requested.key != "energy.restart"
            ):
                checks.append(
                    ValidationCheck(
                        key=check_key,
                        state=CheckState.FAIL,
                        message=(
                            "Ação destrutiva não é automatizada "
                            "pela campanha."
                        ),
                    )
                )
                continue

            context = ExecutionPolicyContext.from_session(session)
            try:
                plan = self.engine.plan(
                    spec.target,
                    requested.key,
                    requested.parameters,
                    context=context,
                )
            except Exception as exc:
                checks.append(
                    ValidationCheck(
                        key=check_key,
                        state=CheckState.FAIL,
                        message=(
                            "Falha ao criar plano: "
                            f"{type(exc).__name__}: {exc}"
                        ),
                    )
                )
                continue

            if not plan.allowed:
                checks.append(
                    ValidationCheck(
                        key=check_key,
                        state=CheckState.FAIL,
                        message="Policy bloqueou a ação.",
                        evidence={
                            "failures": [
                                item.message
                                for item in plan.policy.failures
                            ]
                        },
                    )
                )
                continue

            try:
                record = self.engine.execute(
                    spec.target,
                    requested.key,
                    requested.parameters,
                    context=context,
                    operator=operator,
                    reconnect=self._recover,
                )
            except ExecutionBlockedError as exc:
                checks.append(
                    ValidationCheck(
                        key=check_key,
                        state=CheckState.FAIL,
                        message=str(exc),
                    )
                )
                continue
            except Exception as exc:
                checks.append(
                    ValidationCheck(
                        key=check_key,
                        state=CheckState.FAIL,
                        message=(
                            "Execução falhou: "
                            f"{type(exc).__name__}: {exc}"
                        ),
                    )
                )
                continue

            state_name = record.remediation.validation.status.value
            state = {
                "PASS": CheckState.PASS,
                "FAIL": CheckState.FAIL,
                "UNKNOWN": CheckState.UNKNOWN,
            }.get(state_name, CheckState.UNKNOWN)
            evidence: dict[str, Any] = {
                "transport": record.remediation.command_result.transport,
                "duration_ms": record.duration_ms,
                "validation": state_name,
                "rollback_available": record.rollback_available,
            }
            if record.recovery is not None:
                evidence["recovery"] = asdict(record.recovery)

            checks.append(
                ValidationCheck(
                    key=check_key,
                    state=state,
                    message=record.remediation.validation.message,
                    evidence=evidence,
                )
            )

            if requested.rollback_after:
                checks.append(
                    self._rollback_check(
                        spec,
                        requested,
                        record,
                        operator=operator,
                    )
                )

        return checks

    def _rollback_check(
        self,
        spec: EndpointSpec,
        requested: ControlledActionSpec,
        record,
        *,
        operator: str,
    ) -> ValidationCheck:
        key = f"action.{requested.key}.rollback"
        if not record.rollback_available:
            return ValidationCheck(
                key=key,
                state=CheckState.FAIL,
                message=(
                    "Rollback solicitado, mas a ação não é reversível."
                ),
            )
        try:
            session = self.sessions.open(spec.target, refresh=True)
            rollback = self.engine.rollback(
                record,
                context=ExecutionPolicyContext.from_session(session),
                operator=operator,
            )
        except Exception as exc:
            return ValidationCheck(
                key=key,
                state=CheckState.FAIL,
                message=(
                    "Rollback falhou: "
                    f"{type(exc).__name__}: {exc}"
                ),
            )

        rollback_state = {
            "PASS": CheckState.PASS,
            "FAIL": CheckState.FAIL,
            "UNKNOWN": CheckState.UNKNOWN,
        }.get(
            rollback.validation.status.value,
            CheckState.UNKNOWN,
        )
        return ValidationCheck(
            key=key,
            state=rollback_state,
            message=rollback.validation.message,
            evidence={
                "duration_ms": rollback.duration_ms,
                "transport": rollback.result.transport,
            },
        )

    def validate_endpoint(
        self,
        spec: EndpointSpec,
        *,
        allow_mutations: bool = False,
        allow_high_risk: bool = False,
        allow_reboot: bool = False,
        operator: str = "field-validation",
    ) -> EndpointValidationResult:
        correlation_id = uuid.uuid4().hex
        started = _now()
        checks: list[ValidationCheck] = []
        metadata: dict[str, Any] = {}

        try:
            session = self.sessions.open(spec.target, refresh=True)
            metadata = self._metadata(session)
            checks.extend(self._safe_checks(spec, session))
            if spec.actions:
                checks.extend(
                    self._controlled_checks(
                        spec,
                        session,
                        allow_mutations=allow_mutations,
                        allow_high_risk=allow_high_risk,
                        allow_reboot=allow_reboot,
                        operator=operator,
                    )
                )
        except Exception as exc:
            checks.append(
                ValidationCheck(
                    key="session.open",
                    state=CheckState.FAIL,
                    message=f"{type(exc).__name__}: {exc}",
                )
            )

        return EndpointValidationResult(
            alias=spec.alias,
            target_fingerprint=fingerprint_target(spec.target),
            correlation_id=correlation_id,
            started_at=started,
            finished_at=_now(),
            checks=_sanitize_checks(checks, spec.target),
            metadata=_redact_target(metadata, spec.target),
        )


def load_campaign(path: Path) -> tuple[str, list[EndpointSpec]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("A campanha deve ser um objeto JSON.")

    name = str(payload.get("name") or "central-n2-field-validation")
    raw_endpoints = payload.get("endpoints")
    if not isinstance(raw_endpoints, list):
        raise ValueError("'endpoints' deve ser uma lista.")

    endpoints: list[EndpointSpec] = []
    aliases: set[str] = set()
    for raw in raw_endpoints:
        if not isinstance(raw, dict):
            raise ValueError("Cada endpoint deve ser um objeto.")
        alias = str(raw.get("alias") or "").strip()
        target = str(raw.get("target") or "").strip()
        if not alias or not target:
            raise ValueError("Endpoint exige alias e target.")
        if alias.casefold() in aliases:
            raise ValueError(f"Alias duplicado: {alias}")
        aliases.add(alias.casefold())

        raw_actions = raw.get("actions", [])
        if not isinstance(raw_actions, list):
            raise ValueError(
                f"Endpoint {alias}: 'actions' deve ser uma lista."
            )

        actions = tuple(
            ControlledActionSpec(
                key=str(item.get("key") or "").strip(),
                parameters=dict(item.get("parameters") or {}),
                rollback_after=bool(
                    item.get("rollback_after", False)
                ),
            )
            for item in raw_actions
            if isinstance(item, dict)
        )
        endpoints.append(
            EndpointSpec(
                alias=alias,
                target=target,
                enabled=bool(raw.get("enabled", True)),
                expected_state=(
                    str(raw["expected_state"])
                    if raw.get("expected_state")
                    else None
                ),
                expected_transport=(
                    str(raw["expected_transport"])
                    if raw.get("expected_transport")
                    else None
                ),
                required_capabilities=tuple(
                    str(item)
                    for item in raw.get(
                        "required_capabilities",
                        [],
                    )
                ),
                roles=tuple(
                    str(item)
                    for item in raw.get("roles", [])
                ),
                actions=actions,
            )
        )

    return name, endpoints


def run_campaign(
    runner: EndpointValidationRunner,
    name: str,
    endpoints: list[EndpointSpec],
    **kwargs: Any,
) -> CampaignResult:
    started = _now()
    results = [
        runner.validate_endpoint(endpoint, **kwargs)
        for endpoint in endpoints
        if endpoint.enabled
    ]
    return CampaignResult(
        name=name,
        started_at=started,
        finished_at=_now(),
        endpoints=results,
    )


def write_reports(
    campaign: CampaignResult,
    output_dir: Path,
) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    json_path = output_dir / f"endpoint-validation-{stamp}.json"
    md_path = output_dir / f"endpoint-validation-{stamp}.md"

    json_path.write_text(
        json.dumps(
            campaign.public_dict(),
            indent=2,
            ensure_ascii=False,
            default=str,
        ),
        encoding="utf-8",
    )

    lines = [
        f"# Homologação de endpoints — {campaign.name}",
        "",
        f"Status geral: **{campaign.status.value}**",
        f"Início: {campaign.started_at}",
        f"Fim: {campaign.finished_at}",
        "",
    ]
    for endpoint in campaign.endpoints:
        lines.extend(
            [
                f"## {endpoint.alias}",
                "",
                f"- Status: **{endpoint.status.value}**",
                (
                    "- Target fingerprint: "
                    f"{endpoint.target_fingerprint}"
                ),
                (
                    "- Correlation ID: "
                    f"{endpoint.correlation_id}"
                ),
                "",
                "| Check | Estado | Mensagem |",
                "|---|---|---|",
            ]
        )
        for check in endpoint.checks:
            message = check.message.replace("|", "\\|").replace(
                "\n",
                " ",
            )
            lines.append(
                f"| {check.key} | **{check.state.value}** | "
                f"{message} |"
            )
        lines.append("")

    md_path.write_text("\n".join(lines), encoding="utf-8")
    return json_path, md_path
