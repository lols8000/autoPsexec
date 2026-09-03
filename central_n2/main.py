from __future__ import annotations

import ctypes
import json
import os
import sys
from pathlib import Path

from core.executor import RemoteExecutor
from core.logger import AuditLogger
from ui.console_v3 import ConsoleUIV3


BASE_DIR = Path(__file__).resolve().parent
SETTINGS_PATH = BASE_DIR / "config" / "settings.json"
LOG_DIR = BASE_DIR / "logs"


def is_admin() -> bool:
    if os.name != "nt":
        return True
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except OSError:
        return False


def relaunch_as_admin() -> None:
    params = " ".join(f'"{arg}"' if " " in arg else arg for arg in sys.argv)
    rc = ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, str(BASE_DIR), 1)
    if rc <= 32:
        raise RuntimeError(f"Não foi possível elevar privilégios. ShellExecute retornou {rc}.")


def main() -> int:
    if os.name != "nt":
        print("Esta ferramenta foi projetada para administração de estações Windows.")
        return 2
    if not SETTINGS_PATH.exists():
        print(f"Arquivo de configuração não encontrado: {SETTINGS_PATH}")
        return 2
    if not is_admin():
        try:
            relaunch_as_admin()
        except Exception as exc:
            print(f"Falha ao solicitar elevação administrativa: {exc}")
            return 5
        return 0

    settings = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    logger = AuditLogger(LOG_DIR)
    executor = RemoteExecutor(
        psexec_path=settings.get("psexec_path") or None,
        timeout=int(settings.get("timeout_seconds", 60)),
        logger=logger,
    )
    ConsoleUIV3(executor, SETTINGS_PATH).run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
