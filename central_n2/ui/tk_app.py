from __future__ import annotations

import json
import threading
import tkinter as tk
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any, Callable

from core.config import ConfigLoader
from core.jobs import JobManager, OperationClass
from core.session import SessionManager
from core.validation import validate_host
from modules.disk import DiskModule
from modules.health import HealthModule
from modules.performance import PerformanceModule
from modules.startup import StartupModule
from modules.system import SystemModule
from playbooks import PlaybookRunner, builtin_playbooks


class CentralN2TkApp:
    """GUI leve da Central N2; operações passam pelo mesmo modelo de jobs."""

    def __init__(
        self,
        root: tk.Tk,
        executor,
        settings_path: Path,
        *,
        settings: dict[str, Any] | None = None,
    ) -> None:
        self.root = root
        self.executor = executor
        self.settings = settings or ConfigLoader(settings_path).settings

        runtime = self.settings.get("runtime", {})
        ui = self.settings.get("ui", {})
        self.jobs = JobManager(
            max_workers=int(runtime.get("max_workers", 6)),
            heartbeat_seconds=float(ui.get("heartbeat_seconds", 0.2)),
        )

        self.sessions = SessionManager(executor)
        self.health = HealthModule(executor)
        self.playbooks = builtin_playbooks()
        self.playbook_runner = PlaybookRunner()
        self.performance = PerformanceModule(executor)
        self.disk = DiskModule(executor)
        self.system = SystemModule(executor)
        self.startup = StartupModule(executor)

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.close)

    def _build_ui(self) -> None:
        self.root.title("Central N2 Workstation")
        self.root.geometry("900x620")

        top = tk.Frame(self.root)
        top.pack(fill="x", padx=10, pady=10)

        tk.Label(top, text="Estação:").pack(side="left")
        self.host = tk.StringVar(value="localhost")
        tk.Entry(
            top,
            textvariable=self.host,
            width=35,
        ).pack(side="left", padx=5)

        actions = (
            ("Conectar", self.connect),
            ("Saúde", self.check_health),
            ("Playbook Lentidão", self.slow_playbook),
        )
        for title, callback in actions:
            tk.Button(
                top,
                text=title,
                command=callback,
            ).pack(side="left", padx=4)

        self.status = tk.Label(
            self.root,
            text="Pronto",
            anchor="w",
        )
        self.status.pack(fill="x", padx=10)

        self.output = tk.Text(self.root, wrap="word")
        self.output.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10,
        )

    @staticmethod
    def _serialize(value: Any) -> str:
        if is_dataclass(value):
            value = asdict(value)
        if isinstance(value, str):
            return value
        return json.dumps(
            value,
            ensure_ascii=False,
            indent=2,
            default=str,
        )

    def _target(self) -> str:
        return validate_host(self.host.get().strip())

    def _run(
        self,
        label: str,
        function: Callable[[], Any],
        *,
        operation_class: OperationClass = OperationClass.READ_ONLY,
    ) -> None:
        try:
            host = self._target()
        except ValueError as exc:
            self._finish(label, f"Entrada inválida: {exc}", failed=True)
            return

        self.status.config(text=f"{label} — executando...")
        _, future = self.jobs.submit(
            host,
            label,
            function,
            operation_class=operation_class,
        )

        def waiter() -> None:
            try:
                value = future.result()
                output = self._serialize(value)
                failed = False
            except Exception as exc:
                output = f"ERRO: {type(exc).__name__}: {exc}"
                failed = True

            self.root.after(
                0,
                lambda: self._finish(
                    label,
                    output,
                    failed=failed,
                ),
            )

        threading.Thread(
            target=waiter,
            name="central-n2-gui-waiter",
            daemon=True,
        ).start()

    def _finish(
        self,
        label: str,
        text: str,
        *,
        failed: bool = False,
    ) -> None:
        self.output.delete("1.0", "end")
        self.output.insert("end", text)
        state = "falhou" if failed else "concluído"
        self.status.config(text=f"{label} — {state}")

    def connect(self) -> None:
        host = self._target()
        self._run(
            "Preflight",
            lambda: self.sessions.open(host, refresh=True),
        )

    def check_health(self) -> None:
        host = self._target()

        def collect() -> Any:
            result = self.health.snapshot(host)
            return result.data if result.success else result.stderr

        self._run("Saúde", collect)

    def slow_playbook(self) -> None:
        host = self._target()
        spec = self.playbooks["slow"]
        collectors = {
            "health": self.health.snapshot,
            "performance": lambda target: self.performance.snapshot(
                target,
                8,
                1,
            ),
            "disk": self.disk.space,
            "processes": self.system.list_processes,
            "startup": self.startup.overview,
        }
        self._run(
            "Playbook Lentidão",
            lambda: self.playbook_runner.run(
                spec,
                host,
                collectors,
            ),
            operation_class=OperationClass.HEAVY_READ,
        )

    def close(self) -> None:
        self.jobs.shutdown()
        self.root.destroy()


def run_gui(
    executor,
    settings_path: Path,
    *,
    settings: dict[str, Any] | None = None,
) -> None:
    root = tk.Tk()
    CentralN2TkApp(
        root,
        executor,
        settings_path,
        settings=settings,
    )
    root.mainloop()
