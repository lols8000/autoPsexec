from __future__ import annotations

import getpass
import json
import re
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from time import perf_counter
from typing import Any

from core.redaction import redact, redact_text
from core.result import CommandResult
from core.version import __version__
from execution import DisconnectMode, ExecutionBlockedError
from remediation import ValidationStatus

from .matrix import BASELINE_CASES, PROFILES
from .models import (
    CaseStatus,
    QualificationCase,
    QualificationCaseResult,
    QualificationReport,
)
from .runtime import QualificationRuntime


class QualificationError(RuntimeError):
    pass


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _safe_name(value: str) -> str:
    normalized = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return normalized.strip("._-") or "endpoint"


def _command_evidence(result: CommandResult) -> dict[str, Any]:
    evidence: dict[str, Any] = {
        "success": result.success,
        "transport": result.transport,
        "return_code": result.return_code,
        "duration_ms": result.duration_ms,
        "indeterminate": result.indeterminate,
    }
    if result.data is not None:
        evidence["data"] = result.data
    if result.stderr:
        evidence["stderr"] = result.stderr[:2000]
    if result.stdout and result.data is None:
        evidence["stdout"] = result.stdout[:2000]
    return redact(evidence)


class QualificationRunner:
    def __init__(
        self,
        runtime: QualificationRuntime,
        *,
        output_dir: Path,
        operator: str | None = None,
    ) -> None:
        self.runtime = runtime
        self.output_dir = output_dir
        self.operator = operator or getpass.getuser()

    @staticmethod
    def _case_result(
        case: QualificationCase,
        *,
        status: CaseStatus,
        started_at: str,
        started_clock: float,
        transport: str | None,
        message: str,
        evidence: Any = None,
    ) -> QualificationCaseResult:
        return QualificationCaseResult(
            key=case.key,
            title=case.title,
            category=case.category,
            status=status,
            started_at=started_at,
            finished_at=_utc_now(),
            duration_ms=int((perf_counter() - started_clock) * 1000),
            transport=transport,
            message=message,
            evidence=evidence,
        )

    @staticmethod
    def _profile_expectation(profile: str, session) -> tuple[CaseStatus, str]:
        if profile == "auto":
            return CaseStatus.PASS, "Perfil automático; transporte detectado dinamicamente."

        expected = {
            "local": "local",
            "winrm": "winrm",
            "psexec": "psexec",
        }.get(profile)
        if expected is None:
            return CaseStatus.PASS, f"Perfil {profile} não impõe transporte específico."

        actual = str(session.transport or "unknown").casefold()
        if actual == expected:
            return CaseStatus.PASS, f"Transporte esperado confirmado: {expected}."
        return (
            CaseStatus.FAIL,
            f"Perfil exige {expected}, mas o transporte selecionado foi {actual}.",
        )

    def _should_skip(self, case: QualificationCase, session) -> str | None:
        if case.domain_member and session.capabilities.get("DomainMember") is not True:
            return "Endpoint não é membro de domínio."
        if case.capability and not bool(session.capabilities.get(case.capability)):
            return f"Capability {case.capability} não disponível."
        return None

    def _run_probe(self, case: QualificationCase, host: str, session) -> QualificationCaseResult:
        started_at = _utc_now()
        started = perf_counter()

        skip_reason = self._should_skip(case, session)
        if skip_reason:
            return self._case_result(
                case,
                status=CaseStatus.SKIP,
                started_at=started_at,
                started_clock=started,
                transport=session.transport,
                message=skip_reason,
            )

        handler = self.runtime.probe_handlers().get(case.key)
        if handler is None:
            return self._case_result(
                case,
                status=CaseStatus.UNKNOWN,
                started_at=started_at,
                started_clock=started,
                transport=session.transport,
                message="Probe não registrado no runtime de homologação.",
            )

        try:
            result = handler(host)
        except Exception as exc:
            return self._case_result(
                case,
                status=CaseStatus.FAIL,
                started_at=started_at,
                started_clock=started,
                transport=session.transport,
                message=f"{type(exc).__name__}: {exc}",
            )

        if not isinstance(result, CommandResult):
            return self._case_result(
                case,
                status=CaseStatus.UNKNOWN,
                started_at=started_at,
                started_clock=started,
                transport=session.transport,
                message="Probe retornou contrato inesperado.",
                evidence=result,
            )

        if result.indeterminate:
            status = CaseStatus.UNKNOWN
            message = "Resultado indeterminado; estado final não foi comprovado."
        elif result.success:
            status = CaseStatus.PASS
            message = "Probe concluído com sucesso."
        else:
            status = CaseStatus.FAIL
            message = result.stderr or "Probe falhou."

        return self._case_result(
            case,
            status=status,
            started_at=started_at,
            started_clock=started,
            transport=result.transport or session.transport,
            message=message,
            evidence=_command_evidence(result),
        )

    def _run_action(
        self,
        host: str,
        action_key: str,
        parameters: dict[str, Any],
        session,
        *,
        allow_disruptive: bool,
        confirmed_host: str | None,
    ) -> QualificationCaseResult:
        started_at = _utc_now()
        started = perf_counter()
        try:
            bound = self.runtime.registry.get(action_key)
        except KeyError:
            case = QualificationCase(
                key=f"action.{action_key}",
                title=action_key,
                category="Ação",
                kind="action",
            )
            return self._case_result(
                case,
                status=CaseStatus.FAIL,
                started_at=started_at,
                started_clock=started,
                transport=session.transport,
                message=f"Ação de homologação não encontrada: {action_key}.",
            )

        spec = bound.spec
        case = QualificationCase(
            key=f"action.{action_key}",
            title=spec.title,
            category=f"Ação / {spec.category_label}",
            kind="action",
        )

        disruptive = (
            spec.may_break_connectivity
            or spec.disconnect_mode is not DisconnectMode.NONE
        )
        if disruptive:
            if not allow_disruptive:
                return self._case_result(
                    case,
                    status=CaseStatus.FAIL,
                    started_at=started_at,
                    started_clock=started,
                    transport=session.transport,
                    message=(
                        "Ação disruptiva não autorizada. Use autorização explícita "
                        "de homologação para este endpoint."
                    ),
                )
            if not confirmed_host or confirmed_host.casefold() != host.casefold():
                return self._case_result(
                    case,
                    status=CaseStatus.FAIL,
                    started_at=started_at,
                    started_clock=started,
                    transport=session.transport,
                    message="Confirmação nominal do host ausente ou divergente.",
                )

        try:
            record = self.runtime.engine.execute(
                host,
                action_key,
                parameters,
                context=self.runtime.policy_context(session),
                operator=self.operator,
                reconnect=self.runtime.recover,
            )
        except ExecutionBlockedError as exc:
            return self._case_result(
                case,
                status=CaseStatus.FAIL,
                started_at=started_at,
                started_clock=started,
                transport=session.transport,
                message=f"Policy bloqueou a ação: {exc}",
            )
        except Exception as exc:
            return self._case_result(
                case,
                status=CaseStatus.FAIL,
                started_at=started_at,
                started_clock=started,
                transport=session.transport,
                message=f"{type(exc).__name__}: {exc}",
            )

        validation = record.remediation.validation
        status = {
            ValidationStatus.PASS: CaseStatus.PASS,
            ValidationStatus.FAIL: CaseStatus.FAIL,
            ValidationStatus.UNKNOWN: CaseStatus.UNKNOWN,
        }.get(validation.status, CaseStatus.UNKNOWN)

        evidence = redact(
            {
                "action": spec.key,
                "risk": spec.risk.value,
                "parameters": record.public_parameters,
                "validation": {
                    "status": validation.status.value,
                    "message": validation.message,
                    "evidence": validation.evidence,
                },
                "transport": record.remediation.command_result.transport,
                "duration_ms": record.duration_ms,
                "recovery": (
                    asdict(record.recovery)
                    if record.recovery
                    else None
                ),
                "rollback_available": record.rollback_available,
            }
        )
        return self._case_result(
            case,
            status=status,
            started_at=started_at,
            started_clock=started,
            transport=record.remediation.command_result.transport,
            message=validation.message,
            evidence=evidence,
        )

    def run(
        self,
        host: str,
        *,
        profile: str = "auto",
        requested_actions: list[str] | None = None,
        action_parameters: dict[str, dict[str, Any]] | None = None,
        allow_disruptive: bool = False,
        confirmed_host: str | None = None,
    ) -> QualificationReport:
        if profile not in PROFILES:
            raise QualificationError(f"Perfil de homologação inválido: {profile}")

        requested_actions = requested_actions or []
        action_parameters = action_parameters or {}
        report_started = _utc_now()

        try:
            session = self.runtime.open_session(host, refresh=True)
        except Exception as exc:
            now = _utc_now()
            case = QualificationCase(
                "core.connectivity",
                "Conectividade e transporte administrativo",
                "Core",
                "session",
            )
            failure = QualificationCaseResult(
                key=case.key,
                title=case.title,
                category=case.category,
                status=CaseStatus.FAIL,
                started_at=report_started,
                finished_at=now,
                duration_ms=0,
                message=f"{type(exc).__name__}: {exc}",
            )
            report = QualificationReport(
                host=host,
                profile=profile,
                started_at=report_started,
                finished_at=now,
                transport="unknown",
                connectivity_state="ERROR",
                cases=[failure],
                requested_actions=requested_actions,
                operator=self.operator,
                version=__version__,
            )
            self.export(report)
            return report

        cases: list[QualificationCaseResult] = []
        connectivity_case = BASELINE_CASES[0]
        state = str(session.connectivity.get("state") or "UNKNOWN")
        cases.append(
            QualificationCaseResult(
                key=connectivity_case.key,
                title=connectivity_case.title,
                category=connectivity_case.category,
                status=(CaseStatus.PASS if session.ready else CaseStatus.FAIL),
                started_at=report_started,
                finished_at=_utc_now(),
                duration_ms=0,
                transport=session.transport,
                message=str(session.connectivity.get("diagnosis") or state),
                evidence=session.connectivity,
            )
        )

        capabilities_case = BASELINE_CASES[1]
        cases.append(
            QualificationCaseResult(
                key=capabilities_case.key,
                title=capabilities_case.title,
                category=capabilities_case.category,
                status=(
                    CaseStatus.PASS
                    if session.ready and not session.capability_error
                    else CaseStatus.UNKNOWN
                ),
                started_at=report_started,
                finished_at=_utc_now(),
                duration_ms=0,
                transport=session.transport,
                message=(
                    session.capability_error
                    or "Capabilities coletadas com sucesso."
                ),
                evidence=session.capabilities,
            )
        )

        profile_status, profile_message = self._profile_expectation(profile, session)
        cases.append(
            QualificationCaseResult(
                key="core.profile_expectation",
                title="Compatibilidade com o perfil de homologação",
                category="Core",
                status=profile_status,
                started_at=report_started,
                finished_at=_utc_now(),
                duration_ms=0,
                transport=session.transport,
                message=profile_message,
            )
        )

        if session.ready:
            for case in BASELINE_CASES[2:]:
                cases.append(self._run_probe(case, host, session))
        else:
            for case in BASELINE_CASES[2:]:
                cases.append(
                    QualificationCaseResult(
                        key=case.key,
                        title=case.title,
                        category=case.category,
                        status=CaseStatus.SKIP,
                        started_at=report_started,
                        finished_at=_utc_now(),
                        duration_ms=0,
                        transport=session.transport,
                        message="Sem transporte administrativo utilizável.",
                    )
                )

        for action_key in requested_actions:
            cases.append(
                self._run_action(
                    host,
                    action_key,
                    action_parameters.get(action_key, {}),
                    session,
                    allow_disruptive=allow_disruptive,
                    confirmed_host=confirmed_host,
                )
            )
            refreshed = self.runtime.sessions.get(host)
            if refreshed is not None:
                session = refreshed

        report = QualificationReport(
            host=host,
            profile=profile,
            started_at=report_started,
            finished_at=_utc_now(),
            transport=session.transport,
            connectivity_state=str(
                session.connectivity.get("state") or "UNKNOWN"
            ),
            cases=cases,
            requested_actions=list(requested_actions),
            operator=self.operator,
            version=__version__,
        )
        self.export(report)
        return report

    def export(self, report: QualificationReport) -> tuple[Path, Path]:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        base = f"qualification-{_safe_name(report.host)}-{stamp}"
        json_path = self.output_dir / f"{base}.json"
        md_path = self.output_dir / f"{base}.md"

        payload = redact(asdict(report))
        json_path.write_text(
            json.dumps(
                payload,
                indent=2,
                ensure_ascii=False,
                default=str,
            ),
            encoding="utf-8",
        )

        counts = report.summary()
        lines = [
            f"# Homologação Central N2 — {report.host}",
            "",
            f"- Versão: {report.version or '-'}",
            f"- Perfil: {report.profile}",
            f"- Operador: {report.operator or '-'}",
            f"- Transporte final: {report.transport}",
            f"- Estado: {report.connectivity_state}",
            f"- Início: {report.started_at}",
            f"- Fim: {report.finished_at}",
            "",
            "## Resumo",
            "",
            f"PASS: {counts['PASS']} | FAIL: {counts['FAIL']} | "
            f"UNKNOWN: {counts['UNKNOWN']} | SKIP: {counts['SKIP']}",
            "",
            "## Casos",
            "",
            "| Status | Caso | Categoria | Transporte | Duração | Mensagem |",
            "|---|---|---|---|---:|---|",
        ]
        for case in report.cases:
            message = redact_text(case.message).replace("|", "\\|").replace("\n", " ")
            lines.append(
                f"| {case.status.value} | {case.title} | {case.category} | "
                f"{case.transport or '-'} | {case.duration_ms} ms | {message} |"
            )

        md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return json_path, md_path
