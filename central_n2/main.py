from __future__ import annotations

import argparse
import ctypes
import os
import subprocess
import sys
from pathlib import Path
from typing import Any

from core.config import ConfigError, ConfigLoader
from core.executor import RemoteExecutor
from core.logger import AuditLogger
from core.version import __version__
from execution.config_validation import validate_execution_configuration

SOURCE_DIR = Path(__file__).resolve().parent
BASE_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else SOURCE_DIR
)
SETTINGS_PATH = BASE_DIR / "config" / "settings.json"
if not SETTINGS_PATH.exists():
    SETTINGS_PATH = SOURCE_DIR / "config" / "settings.json"
LOG_DIR = BASE_DIR / "logs"


def is_admin() -> bool:
    if os.name != "nt":
        return True
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except (AttributeError, OSError):
        return False


def relaunch_as_admin() -> None:
    if getattr(sys, "frozen", False):
        arguments = sys.argv[1:]
    else:
        arguments = [str(SOURCE_DIR / "main.py"), *sys.argv[1:]]

    parameters = subprocess.list2cmdline(arguments)
    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        sys.executable,
        parameters,
        str(BASE_DIR),
        1,
    )
    if result <= 32:
        raise RuntimeError(
            "Não foi possível elevar privilégios. "
            f"ShellExecute retornou {result}."
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Central N2 Workstation"
    )
    parser.add_argument("--gui", action="store_true")
    parser.add_argument("--version", action="store_true")
    return parser.parse_args()


def build_executor(
    settings: dict[str, Any],
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

    if args.version:
        print(__version__)
        return 0

    if os.name != "nt":
        print(
            "Esta ferramenta foi projetada para administração "
            "de estações Windows."
        )
        return 2

    if not SETTINGS_PATH.exists():
        print(
            f"Arquivo de configuração não encontrado: {SETTINGS_PATH}"
        )
        return 2

    if not is_admin():
        try:
            relaunch_as_admin()
        except Exception as exc:
            print(
                "Falha ao solicitar elevação administrativa: "
                f"{exc}"
            )
            return 5
        return 0

    try:
        settings = ConfigLoader(SETTINGS_PATH).settings
    except ConfigError as exc:
        print(f"Configuração inválida: {exc}")
        return 2

    catalog_report = validate_execution_configuration(settings)
    settings = catalog_report.settings

    logging_config = settings.get("logging", {})
    logger = AuditLogger(
        LOG_DIR,
        verbose_payloads=bool(
            logging_config.get("verbose_payloads", False)
        ),
        max_error_chars=int(
            logging_config.get("max_error_chars", 4000)
        ),
    )
    if catalog_report.issues:
        print("\n⚠ Catálogo de execuções com entradas desabilitadas:")
        for issue in catalog_report.issues:
            print(
                f" - {issue.section}.{issue.key}: {issue.message}"
            )
        logger.log_event(
            "execution_catalog_validation",
            "local-ui",
            "warning",
            issues=[
                {
                    "section": issue.section,
                    "key": issue.key,
                    "message": issue.message,
                }
                for issue in catalog_report.issues
            ],
            enabled_counts=catalog_report.enabled_counts,
        )

    executor = build_executor(settings, logger)

    if args.gui:
        from ui.tk_app import run_gui

        run_gui(
            executor,
            SETTINGS_PATH,
            settings=settings,
        )
    else:
        from ui.console_v5 import ConsoleUIV5

        ConsoleUIV5(
            executor,
            SETTINGS_PATH,
            settings=settings,
        ).run()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
