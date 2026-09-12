from __future__ import annotations

from pathlib import Path
from typing import Any

from execution.config_validation import validate_execution_configuration

from .matrix import PROFILES
from .runner import QualificationError, QualificationRunner
from .runtime import QualificationRuntime


def parse_action_parameters(values: list[str] | None) -> dict[str, dict[str, Any]]:
    parsed: dict[str, dict[str, Any]] = {}
    for raw in values or []:
        if ":" not in raw or "=" not in raw:
            raise QualificationError(
                "Parâmetro de homologação deve usar ACTION:KEY=VALUE."
            )
        action, remainder = raw.split(":", 1)
        key, value = remainder.split("=", 1)
        action = action.strip()
        key = key.strip()
        if not action or not key:
            raise QualificationError(
                "Parâmetro de homologação deve usar ACTION:KEY=VALUE."
            )
        parsed.setdefault(action, {})[key] = value
    return parsed


def run_qualification_cli(
    *,
    executor,
    settings_path: Path,
    settings: dict[str, Any],
    hosts: list[str],
    profile: str,
    actions: list[str] | None,
    raw_parameters: list[str] | None,
    output_dir: Path,
    allow_disruptive: bool,
    confirmed_host: str | None,
) -> int:
    if profile not in PROFILES:
        print(f"Perfil inválido: {profile}")
        return 2
    if not hosts:
        print("Informe ao menos um endpoint para homologação.")
        return 2
    if allow_disruptive and len(hosts) != 1:
        print(
            "Ações disruptivas exigem exatamente um host por execução "
            "para manter confirmação nominal inequívoca."
        )
        return 2

    try:
        parameters = parse_action_parameters(raw_parameters)
    except QualificationError as exc:
        print(f"Configuração da homologação inválida: {exc}")
        return 2

    catalog_report = validate_execution_configuration(settings)
    runtime = QualificationRuntime(
        executor,
        settings_path,
        catalog_report.settings,
    )
    runner = QualificationRunner(
        runtime,
        output_dir=output_dir,
    )

    failed = False
    for host in hosts:
        print(f"\n=== HOMOLOGAÇÃO: {host} ===")
        try:
            report = runner.run(
                host,
                profile=profile,
                requested_actions=list(actions or []),
                action_parameters=parameters,
                allow_disruptive=allow_disruptive,
                confirmed_host=confirmed_host,
            )
        except (QualificationError, KeyError, ValueError) as exc:
            print(f"✗ {type(exc).__name__}: {exc}")
            failed = True
            continue

        summary = report.summary()
        print(
            f"Transporte: {report.transport} | "
            f"Estado: {report.connectivity_state}"
        )
        print(
            f"PASS={summary['PASS']} FAIL={summary['FAIL']} "
            f"UNKNOWN={summary['UNKNOWN']} SKIP={summary['SKIP']}"
        )
        print(f"Evidências: {output_dir}")
        if not report.passed:
            failed = True

    return 4 if failed else 0
