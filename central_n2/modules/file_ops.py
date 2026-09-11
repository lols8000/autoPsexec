from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import (
    quote_powershell_literal,
    validate_windows_path,
)


class FileOperationsModule:
    """Operações explícitas de arquivo; não expõe shell livre."""

    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def ensure_directory(self, host: str, path: str) -> CommandResult:
        safe_path = quote_powershell_literal(
            validate_windows_path(path)
        )
        script = f"""
$item = New-Item -ItemType Directory -Path {safe_path} -Force -ErrorAction Stop
[pscustomobject]@{{Exists=(Test-Path {safe_path});Path=$item.FullName}}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=120,
        )

    def path_status(
        self,
        host: str,
        source: str,
        destination: str,
    ) -> CommandResult:
        safe_source = quote_powershell_literal(
            validate_windows_path(source)
        )
        safe_destination = quote_powershell_literal(
            validate_windows_path(destination)
        )
        script = f"""
[pscustomobject]@{{
    Source = {safe_source}
    Destination = {safe_destination}
    SourceExists = [bool](Test-Path -LiteralPath {safe_source})
    DestinationExists = [bool](Test-Path -LiteralPath {safe_destination})
}}
"""
        return self.executor.execute_powershell_json(
            host,
            script,
            timeout=90,
        )

    def move_path(
        self,
        host: str,
        source: str,
        destination: str,
    ) -> CommandResult:
        safe_source = quote_powershell_literal(
            validate_windows_path(source)
        )
        safe_destination = quote_powershell_literal(
            validate_windows_path(destination)
        )
        script = f"""
if (-not (Test-Path -LiteralPath {safe_source})) {{
    throw 'Origem não encontrada.'
}}
Move-Item -LiteralPath {safe_source} -Destination {safe_destination} -ErrorAction Stop
[pscustomobject]@{{
    SourceExists = (Test-Path -LiteralPath {safe_source})
    DestinationExists = (Test-Path -LiteralPath {safe_destination})
    Destination = {safe_destination}
}}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=300,
        )

    def remove_file(self, host: str, path: str) -> CommandResult:
        safe_path = quote_powershell_literal(
            validate_windows_path(path)
        )
        script = f"""
if (-not (Test-Path -LiteralPath {safe_path} -PathType Leaf)) {{
    throw 'Arquivo não encontrado ou não é um arquivo.'
}}
Remove-Item -LiteralPath {safe_path} -Force -ErrorAction Stop
[pscustomobject]@{{Removed=[bool](-not (Test-Path -LiteralPath {safe_path}));Path={safe_path}}}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=120,
        )
