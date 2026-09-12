from __future__ import annotations

from typing import Any

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import (
    quote_powershell_literal,
    validate_safe_name,
)


class RegistryActionsModule:
    """Executa somente alterações de Registro previamente configuradas."""

    ALLOWED_TYPES = {
        "String": "String",
        "ExpandString": "ExpandString",
        "DWord": "DWord",
        "QWord": "QWord",
        "MultiString": "MultiString",
        "Binary": "Binary",
    }

    def __init__(
        self,
        executor: RemoteExecutor,
        settings: dict[str, Any],
    ) -> None:
        self.executor = executor
        self.catalog = settings.get("registry_actions", {})

    def keys(self) -> tuple[str, ...]:
        return tuple(sorted(self.catalog))

    def _definition(
        self,
        host: str,
        key: str,
    ) -> tuple[str, dict[str, Any]] | CommandResult:
        safe_key = validate_safe_name(
            key,
            label="Ação de Registro",
        )
        item = self.catalog.get(safe_key)
        if not isinstance(item, dict):
            return CommandResult.failure(
                host,
                safe_key,
                "Ação de Registro homologada não encontrada.",
            )

        path = str(item.get("path") or "")
        if not path.upper().startswith("HKLM:\\"):
            return CommandResult.failure(
                host,
                safe_key,
                "Somente HKLM é permitido em ações homologadas.",
            )

        name = str(item.get("name") or "")
        if not name:
            return CommandResult.failure(
                host,
                safe_key,
                "Nome do valor não configurado.",
            )

        return safe_key, item

    @staticmethod
    def _render_value(
        value: Any,
        value_type: str | None = None,
    ) -> str:
        if isinstance(value, str):
            return quote_powershell_literal(value)
        if isinstance(value, bool):
            return "$true" if value else "$false"
        if isinstance(value, (int, float)):
            return str(value)
        if isinstance(value, list):
            if value_type == "Binary":
                try:
                    numbers = ",".join(
                        str(int(entry))
                        for entry in value
                    )
                except (TypeError, ValueError) as exc:
                    raise ValueError(
                        "Valor Binary deve conter apenas bytes numéricos."
                    ) from exc
                return f"@({numbers})"
            items = ",".join(
                quote_powershell_literal(str(entry))
                for entry in value
            )
            return f"@({items})"
        raise ValueError("Tipo de valor não suportado.")

    def inspect(self, host: str, key: str) -> CommandResult:
        definition = self._definition(host, key)
        if isinstance(definition, CommandResult):
            return definition

        safe_key, item = definition
        path = str(item["path"])
        name = str(item["name"])
        safe_path = quote_powershell_literal(path)
        safe_name = quote_powershell_literal(name)

        script = f"""
$exists = $false
$value = $null
$valueKind = $null
if (Test-Path {safe_path}) {{
    try {{
        $key = Get-Item -Path {safe_path} -ErrorAction Stop
        $value = Get-ItemPropertyValue -Path {safe_path} -Name {safe_name} -ErrorAction Stop
        $valueKind = $key.GetValueKind({safe_name}).ToString()
        $exists = $true
    }} catch {{}}
}}
[pscustomobject]@{{
    Exists = [bool]$exists
    Value = $value
    ValueKind = $valueKind
    Path = {safe_path}
    Name = {safe_name}
    CatalogKey = '{safe_key}'
}}
"""
        return self.executor.execute_powershell_json(
            host,
            script,
            timeout=90,
        )

    def apply(self, host: str, key: str) -> CommandResult:
        definition = self._definition(host, key)
        if isinstance(definition, CommandResult):
            return definition

        safe_key, item = definition
        path = str(item["path"])
        name = str(item["name"])
        mode = str(item.get("mode") or "set").casefold()
        safe_path = quote_powershell_literal(path)
        safe_name = quote_powershell_literal(name)

        if mode == "remove":
            script = f"""
if (Test-Path {safe_path}) {{
    Remove-ItemProperty -Path {safe_path} -Name {safe_name} -ErrorAction SilentlyContinue
}}
[pscustomobject]@{{Applied=$true;Mode='remove';Path={safe_path};Name={safe_name}}}
"""
        elif mode == "set":
            value_type = str(item.get("type") or "String")
            allowed_type = self.ALLOWED_TYPES.get(value_type)
            if not allowed_type:
                return CommandResult.failure(
                    host,
                    safe_key,
                    f"Tipo de Registro não permitido: {value_type}",
                )
            try:
                ps_value = self._render_value(
                    item.get("value"),
                    allowed_type,
                )
            except ValueError as exc:
                return CommandResult.failure(
                    host,
                    safe_key,
                    str(exc),
                )

            script = f"""
New-Item -Path {safe_path} -Force -ErrorAction Stop | Out-Null
New-ItemProperty -Path {safe_path} -Name {safe_name} -Value {ps_value} -PropertyType {allowed_type} -Force -ErrorAction Stop | Out-Null
$value = Get-ItemPropertyValue -Path {safe_path} -Name {safe_name} -ErrorAction Stop
[pscustomobject]@{{Applied=$true;Mode='set';Path={safe_path};Name={safe_name};Value=$value}}
"""
        else:
            return CommandResult.failure(
                host,
                safe_key,
                "mode deve ser set ou remove.",
            )

        result = self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=120,
        )
        result.metadata["registry_action_key"] = safe_key
        return result

    def rollback(
        self,
        host: str,
        key: str,
        before: Any,
    ) -> CommandResult:
        definition = self._definition(host, key)
        if isinstance(definition, CommandResult):
            return definition

        safe_key, item = definition
        evidence = (
            before.data
            if isinstance(before, CommandResult)
            else before
        )
        if not isinstance(evidence, dict) or "Exists" not in evidence:
            return CommandResult.failure(
                host,
                safe_key,
                "Estado anterior do Registro não está disponível.",
            )

        path = str(item["path"])
        name = str(item["name"])
        safe_path = quote_powershell_literal(path)
        safe_name = quote_powershell_literal(name)

        if evidence.get("Exists") is True:
            value_type = str(evidence.get("ValueKind") or "")
            allowed_type = self.ALLOWED_TYPES.get(value_type)
            if not allowed_type:
                return CommandResult.failure(
                    host,
                    safe_key,
                    (
                        "Tipo original do Registro não é restaurável "
                        f"automaticamente: {value_type or 'desconhecido'}"
                    ),
                )
            try:
                previous = self._render_value(
                    evidence.get("Value"),
                    allowed_type,
                )
            except ValueError as exc:
                return CommandResult.failure(
                    host,
                    safe_key,
                    str(exc),
                )
            script = f"""
New-Item -Path {safe_path} -Force -ErrorAction Stop | Out-Null
New-ItemProperty -Path {safe_path} -Name {safe_name} -Value {previous} -PropertyType {allowed_type} -Force -ErrorAction Stop | Out-Null
[pscustomobject]@{{Applied=$true;Rollback=$true;Restored='value';Path={safe_path};Name={safe_name}}}
"""
        else:
            script = f"""
if (Test-Path {safe_path}) {{
    Remove-ItemProperty -Path {safe_path} -Name {safe_name} -ErrorAction SilentlyContinue
}}
[pscustomobject]@{{Applied=$true;Rollback=$true;Restored='absent';Path={safe_path};Name={safe_name}}}
"""

        result = self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=120,
        )
        result.metadata["registry_action_key"] = safe_key
        result.metadata["rollback"] = True
        return result
