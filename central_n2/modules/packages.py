from __future__ import annotations

import shutil
from pathlib import Path, PureWindowsPath
from typing import Any

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import quote_cmd_argument, validate_safe_name


class PackagesModule:
    """Instala pacotes previamente homologados em settings.local.json."""

    def __init__(
        self,
        executor: RemoteExecutor,
        settings: dict[str, Any],
    ) -> None:
        self.executor = executor
        self.catalog = settings.get("packages", {})

    def keys(self) -> tuple[str, ...]:
        return tuple(sorted(self.catalog))

    def _item(
        self,
        host: str,
        key: str,
    ) -> dict[str, Any] | CommandResult:
        item = self.catalog.get(key)
        if not isinstance(item, dict):
            return CommandResult.failure(
                host,
                "package",
                f"Pacote homologado '{key}' não encontrado.",
            )
        return item

    def _copy_destination(
        self,
        host: str,
        remote_path: str,
    ) -> str:
        is_local = getattr(self.executor, "is_local", None)
        if callable(is_local) and is_local(host):
            path = PureWindowsPath(remote_path)
            if not path.drive:
                raise ValueError("remote_path deve ser absoluto.")
            return str(path)
        return self._admin_destination(host, remote_path)

    @staticmethod
    def _admin_destination(host: str, remote_path: str) -> str:
        path = PureWindowsPath(remote_path)
        if not path.drive:
            raise ValueError("remote_path deve ser absoluto.")
        drive = path.drive.rstrip(":")
        relative = "\\".join(path.parts[1:])
        return f"\\\\{host}\\{drive}$\\{relative}"

    def install(self, host: str, key: str) -> CommandResult:
        safe_key = validate_safe_name(key, label="Pacote")
        item = self._item(host, safe_key)
        if isinstance(item, CommandResult):
            return item

        source = str(item.get("source") or "").strip()
        if not source:
            return CommandResult.failure(
                host,
                safe_key,
                "Pacote sem source configurado.",
            )
        source_path = Path(source)
        if not source_path.exists() or not source_path.is_file():
            return CommandResult.failure(
                host,
                safe_key,
                f"Source do pacote não encontrado: {source}",
            )

        installer_type = str(item.get("type") or "").casefold()
        if installer_type not in {"msi", "exe", "cmd", "bat"}:
            return CommandResult.failure(
                host,
                safe_key,
                "Tipo permitido: msi, exe, cmd ou bat.",
            )

        remote_path = str(
            item.get("remote_path")
            or rf"C:\CentralN2\Packages\{source_path.name}"
        )
        try:
            destination = self._copy_destination(host, remote_path)
            Path(destination).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source_path, destination)
        except (OSError, ValueError) as exc:
            return CommandResult.failure(
                host,
                safe_key,
                f"Falha ao copiar pacote: {exc}",
            )

        args = str(item.get("args") or "").strip()
        quoted = quote_cmd_argument(remote_path)
        if installer_type == "msi":
            command = f"msiexec.exe /i {quoted} {args}".strip()
        elif installer_type in {"cmd", "bat"}:
            command = f"cmd.exe /d /c {quoted} {args}".strip()
        else:
            command = f"{quoted} {args}".strip()

        result = self.executor.execute_mutating_cmd(
            host,
            command,
            timeout=int(item.get("timeout_seconds", 1800)),
        )
        result.metadata["package_key"] = safe_key
        result.metadata["remote_installer"] = remote_path
        result.metadata["cleanup_remote_installer"] = bool(
            item.get("cleanup", True)
        )

        if result.success and item.get("cleanup", True):
            cleanup = self.executor.execute_mutating_cmd(
                host,
                f"del /f /q {quote_cmd_argument(remote_path)}",
                timeout=60,
            )
            result.metadata["cleanup_success"] = cleanup.success

        return result
