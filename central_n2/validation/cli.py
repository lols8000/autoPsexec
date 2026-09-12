from __future__ import annotations

import argparse
import getpass
from pathlib import Path

from core.config import ConfigLoader
from core.executor import RemoteExecutor
from core.logger import AuditLogger
from execution.config_validation import validate_execution_configuration

from .models import CheckState
from .runner import (
    EndpointValidationRunner,
    load_campaign,
    run_campaign,
    write_reports,
)


ACKNOWLEDGEMENT = "VALIDAR ENDPOINTS 5.2"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Homologação controlada de endpoints reais "
            "da Central N2 Workstation 5.2"
        )
    )
    parser.add_argument(
        "--matrix",
        type=Path,
        required=True,
        help="Arquivo JSON local com a matriz de endpoints.",
    )
    parser.add_argument(
        "--settings",
        type=Path,
        default=Path("config/settings.json"),
        help="settings.json da Central N2.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("reports/field-validation"),
        help="Diretório local para evidências.",
    )
    parser.add_argument(
        "--allow-mutations",
        action="store_true",
        help="Permite somente ações explicitamente listadas na matriz.",
    )
    parser.add_argument(
        "--allow-high-risk",
        action="store_true",
        help="Permite ações HIGH explicitamente listadas.",
    )
    parser.add_argument(
        "--allow-reboot",
        action="store_true",
        help="Permite somente a ação CRITICAL energy.restart.",
    )
    parser.add_argument(
        "--acknowledge",
        default="",
        help=(
            "Obrigatório quando houver mutações. "
            f"Valor exato: {ACKNOWLEDGEMENT}"
        ),
    )
    return parser.parse_args()


def _validate_flags(args: argparse.Namespace) -> None:
    if args.allow_high_risk and not args.allow_mutations:
        raise ValueError(
            "--allow-high-risk exige --allow-mutations."
        )
    if args.allow_reboot and not args.allow_mutations:
        raise ValueError(
            "--allow-reboot exige --allow-mutations."
        )
    if (
        args.allow_mutations
        and args.acknowledge != ACKNOWLEDGEMENT
    ):
        raise ValueError(
            "Execução mutável exige confirmação explícita com "
            f"--acknowledge \"{ACKNOWLEDGEMENT}\"."
        )


def _build_executor(
    settings: dict,
    logger: AuditLogger,
) -> RemoteExecutor:
    runtime = settings.get("runtime", {})
    return RemoteExecutor(
        psexec_path=settings.get("psexec_path") or None,
        timeout=int(settings.get("timeout_seconds", 60)),
        logger=logger,
        transport_cache_ttl_seconds=float(
            runtime.get("transport_cache_ttl_seconds", 120)
        ),
        retry_attempts=int(runtime.get("retry_attempts", 2)),
        retry_base_delay_seconds=float(
            runtime.get("retry_base_delay_seconds", 0.5)
        ),
    )


def main() -> int:
    args = parse_args()
    try:
        _validate_flags(args)
    except ValueError as exc:
        print(f"Configuração insegura: {exc}")
        return 2

    if not args.matrix.exists():
        print(f"Matriz não encontrada: {args.matrix}")
        return 2
    if not args.settings.exists():
        print(f"Settings não encontrado: {args.settings}")
        return 2

    try:
        settings = ConfigLoader(args.settings).settings
        catalog = validate_execution_configuration(settings)
        settings = catalog.settings
        name, endpoints = load_campaign(args.matrix)
    except Exception as exc:
        print(
            "Falha ao carregar campanha: "
            f"{type(exc).__name__}: {exc}"
        )
        return 2

    enabled = [item for item in endpoints if item.enabled]
    if not enabled:
        print("Nenhum endpoint está habilitado na matriz.")
        return 2

    log_dir = args.output_dir / "logs"
    logger = AuditLogger(log_dir)
    executor = _build_executor(settings, logger)
    runner = EndpointValidationRunner(
        executor,
        settings,
        args.settings,
    )

    print(
        f"Campanha: {name} | endpoints={len(enabled)} | "
        f"modo={'CONTROLLED' if args.allow_mutations else 'SAFE'}"
    )
    if args.allow_mutations:
        print(
            "Mutações habilitadas somente para ações declaradas "
            "explicitamente na matriz."
        )

    campaign = run_campaign(
        runner,
        name,
        endpoints,
        allow_mutations=args.allow_mutations,
        allow_high_risk=args.allow_high_risk,
        allow_reboot=args.allow_reboot,
        operator=getpass.getuser(),
    )
    json_path, md_path = write_reports(
        campaign,
        args.output_dir,
    )

    for endpoint in campaign.endpoints:
        print(
            f"{endpoint.alias}: {endpoint.status.value} "
            f"[{endpoint.target_fingerprint}]"
        )

    print(f"Status geral: {campaign.status.value}")
    print(f"JSON: {json_path}")
    print(f"Markdown: {md_path}")

    if campaign.status is CheckState.PASS:
        return 0
    if campaign.status is CheckState.UNKNOWN:
        return 3
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
