from __future__ import annotations

import shutil
from pathlib import Path, PureWindowsPath
from typing import Any

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import (
    quote_powershell_literal,
    validate_safe_name,
)


class CertificatesModule:
    """Importa somente certificados públicos previamente homologados."""

    ALLOWED_STORES = {"Root", "CA", "My", "TrustedPeople"}

    def __init__(
        self,
        executor: RemoteExecutor,
        settings: dict[str, Any],
    ) -> None:
        self.executor = executor
        self.catalog = settings.get("certificates", {})

    def keys(self) -> tuple[str, ...]:
        return tuple(sorted(self.catalog))

    @staticmethod
    def _admin_destination(host: str, remote_path: str) -> str:
        path = PureWindowsPath(remote_path)
        if not path.drive:
            raise ValueError("remote_path deve ser absoluto.")
        drive = path.drive.rstrip(":")
        relative = "\\".join(path.parts[1:])
        return f"\\\\{host}\\{drive}$\\{relative}"

    def import_certificate(self, host: str, key: str) -> CommandResult:
        safe_key = validate_safe_name(key, label="Certificado")
        item = self.catalog.get(safe_key)
        if not isinstance(item, dict):
            return CommandResult.failure(
                host,
                safe_key,
                "Certificado homologado não encontrado.",
            )

        source = Path(str(item.get("source") or ""))
        if (
            not source.exists()
            or not source.is_file()
            or source.suffix.casefold() not in {".cer", ".crt"}
        ):
            return CommandResult.failure(
                host,
                safe_key,
                "Source deve ser arquivo .cer/.crt existente.",
            )

        store = str(item.get("store") or "Root")
        if store not in self.ALLOWED_STORES:
            return CommandResult.failure(
                host,
                safe_key,
                f"Store não permitido: {store}",
            )

        remote_path = str(
            item.get("remote_path")
            or rf"C:\CentralN2\Certificates\{source.name}"
        )
        try:
            destination = self._admin_destination(host, remote_path)
            Path(destination).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)
        except (OSError, ValueError) as exc:
            return CommandResult.failure(
                host,
                safe_key,
                f"Falha ao copiar certificado: {exc}",
            )

        safe_remote = quote_powershell_literal(remote_path)
        safe_store = quote_powershell_literal(
            rf"Cert:\LocalMachine\{store}"
        )
        script = f"""
$cert = Import-Certificate -FilePath {safe_remote} -CertStoreLocation {safe_store} -ErrorAction Stop
Remove-Item -LiteralPath {safe_remote} -Force -ErrorAction SilentlyContinue
[pscustomobject]@{{
    Imported = [bool]$cert
    Thumbprint = $cert.Thumbprint
    Subject = $cert.Subject
    Store = {safe_store}
}}
"""
        result = self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=180,
        )
        result.metadata["certificate_key"] = safe_key
        return result
