from __future__ import annotations

import time
from pathlib import Path
from typing import Any

from core.session import SessionManager
from execution import (
    ExecutionDependencies,
    ExecutionEngine,
    ExecutionPolicyContext,
    RecoveryResult,
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


class QualificationRuntime:
    """Runtime mínimo para homologação real de endpoints.

    Usa os mesmos módulos, SessionManager, catálogo e ExecutionEngine da Central,
    evitando uma segunda implementação de transporte ou regras de segurança.
    """

    def __init__(
        self,
        executor,
        settings_path: Path,
        settings: dict[str, Any],
    ) -> None:
        self.executor = executor
        self.settings_path = settings_path
        self.settings = settings
        self.sessions = SessionManager(executor)

        self.health = HealthModule(executor)
        self.system = SystemModule(executor)
        self.network = NetworkModule(executor)
        self.devices = DevicesModule(executor)
        self.security = SecurityModule(executor)
        self.domain = DomainModule(executor)
        self.printers = PrintersModule(executor)
        self.updates = UpdatesModule(executor)
        self.users = UsersProfilesModule(executor)
        self.disk = DiskModule(executor)
        self.repair = RepairModule(executor)
        self.software = SoftwareModule(
            executor,
            settings_path,
            settings=settings,
        )
        self.glpi = GLPIModule(
            executor,
            settings_path,
            settings=settings,
        )
        self.packages = PackagesModule(executor, settings)
        self.certificates = CertificatesModule(executor, settings)
        self.registry_actions = RegistryActionsModule(executor, settings)
        execution_settings = settings.get("execution", {})
        self.files = FileOperationsModule(
            executor,
            allowed_roots=execution_settings.get("file_roots"),
        )

        self.registry = build_execution_registry(
            ExecutionDependencies(
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
        )
        self.engine = ExecutionEngine(self.registry)

    def open_session(self, host: str, *, refresh: bool = True):
        return self.sessions.open(host, refresh=refresh)

    @staticmethod
    def policy_context(session) -> ExecutionPolicyContext:
        return ExecutionPolicyContext.from_session(session)

    def recover(
        self,
        host: str,
        timeout_seconds: int,
        delay_seconds: int,
    ) -> RecoveryResult:
        if delay_seconds > 0:
            time.sleep(delay_seconds)

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

            time.sleep(5)

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

    def probe_handlers(self):
        return {
            "core.health": self.health.snapshot,
            "inventory.processes": self.system.list_processes,
            "inventory.services": self.system.list_services,
            "network.adapters": self.network.adapters,
            "network.ip": self.network.ip_configuration,
            "devices.problems": self.devices.problem_devices,
            "security.posture": self.security.status,
            "domain.status": self.domain.status,
            "printers.inventory": self.printers.list_printers,
            "glpi.status": self.glpi.status,
            "updates.status": self.updates.status,
        }
