from __future__ import annotations
import argparse,ctypes,os,sys
from pathlib import Path
from core.config import ConfigLoader
from core.executor import RemoteExecutor
from core.logger import AuditLogger
from core.version import __version__
SOURCE_DIR=Path(__file__).resolve().parent;BASE_DIR=Path(sys.executable).resolve().parent if getattr(sys,"frozen",False) else SOURCE_DIR;SETTINGS_PATH=BASE_DIR/"config"/"settings.json"
if not SETTINGS_PATH.exists():SETTINGS_PATH=SOURCE_DIR/"config"/"settings.json"
LOG_DIR=BASE_DIR/"logs"
def is_admin():
    if os.name!="nt":return True
    try:return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except OSError:return False
def relaunch_as_admin():
    if getattr(sys, "frozen", False):
        args = sys.argv[1:]
    else:
        args = [str(SOURCE_DIR / "main.py"), *sys.argv[1:]]
    params = " ".join(f'"{a}"' if " " in a else a for a in args)
    rc = ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, str(BASE_DIR), 1)
    if rc <= 32:
        raise RuntimeError(f"Não foi possível elevar privilégios. ShellExecute retornou {rc}.")
def parse_args():
    p=argparse.ArgumentParser(description="Central N2 Workstation");p.add_argument("--gui",action="store_true");p.add_argument("--version",action="store_true");return p.parse_args()
def main():
    args=parse_args()
    if args.version:print(__version__);return 0
    if os.name!="nt":print("Esta ferramenta foi projetada para administração de estações Windows.");return 2
    if not SETTINGS_PATH.exists():print(f"Arquivo de configuração não encontrado: {SETTINGS_PATH}");return 2
    if not is_admin():
        try:relaunch_as_admin()
        except Exception as exc:print(f"Falha ao solicitar elevação administrativa: {exc}");return 5
        return 0
    settings=ConfigLoader(SETTINGS_PATH).settings;runtime=settings.get("runtime",{});logger=AuditLogger(LOG_DIR);executor=RemoteExecutor(psexec_path=settings.get("psexec_path") or None,timeout=int(settings.get("timeout_seconds",60)),logger=logger,transport_cache_ttl_seconds=float(runtime.get("transport_cache_ttl_seconds",120)))
    if args.gui:
        from ui.tk_app import run_gui;run_gui(executor,SETTINGS_PATH)
    else:
        from ui.console_v5 import ConsoleUIV5;ConsoleUIV5(executor,SETTINGS_PATH).run()
    return 0
if __name__=="__main__":raise SystemExit(main())
