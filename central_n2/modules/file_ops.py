from __future__ import annotations

from pathlib import PureWindowsPath
from typing import Iterable

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import (
    quote_powershell_literal,
    validate_windows_path,
)


_DEFAULT_FILE_ROOTS = (
    r"C:\CentralN2",
    r"C:\Temp",
)


class FileOperationsModule:
    """Operações explícitas e confinadas a raízes autorizadas."""

    def __init__(
        self,
        executor: RemoteExecutor,
        *,
        allowed_roots: Iterable[str] | None = None,
    ) -> None:
        self.executor = executor
        roots = tuple(allowed_roots or _DEFAULT_FILE_ROOTS)
        if not roots:
            raise ValueError(
                "FileOperationsModule exige ao menos uma raiz permitida."
            )

        normalized: list[str] = []
        for root in roots:
            candidate = PureWindowsPath(
                validate_windows_path(str(root))
            )
            if ".." in candidate.parts:
                raise ValueError(
                    f"Raiz de arquivo não pode conter '..': {root}"
                )
            value = str(candidate)
            if value.casefold() not in {
                item.casefold()
                for item in normalized
            }:
                normalized.append(value)

        self.allowed_roots = tuple(normalized)

    def _validate_scoped_path(self, value: str) -> str:
        raw = validate_windows_path(value)
        candidate = PureWindowsPath(raw)

        if ".." in candidate.parts:
            raise ValueError(
                "Caminho com '..' não é permitido em operações de arquivo."
            )

        candidate_parts = tuple(
            part.casefold()
            for part in candidate.parts
        )

        for root_text in self.allowed_roots:
            root = PureWindowsPath(root_text)
            root_parts = tuple(
                part.casefold()
                for part in root.parts
            )
            if (
                len(candidate_parts) >= len(root_parts)
                and candidate_parts[: len(root_parts)] == root_parts
            ):
                return str(candidate)

        roots = ", ".join(self.allowed_roots)
        raise ValueError(
            f"Caminho fora das raízes permitidas: {roots}."
        )

    def ensure_directory(self, host: str, path: str) -> CommandResult:
        safe_path = quote_powershell_literal(
            self._validate_scoped_path(path)
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
            self._validate_scoped_path(source)
        )
        safe_destination = quote_powershell_literal(
            self._validate_scoped_path(destination)
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
            self._validate_scoped_path(source)
        )
        safe_destination = quote_powershell_literal(
            self._validate_scoped_path(destination)
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
            self._validate_scoped_path(path)
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
