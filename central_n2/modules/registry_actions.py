from __future__ import annotations

from typing import Any

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import quote_powershell_literal, validate_safe_name


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

    def apply(self, host: str, key: str) -> CommandResult:
        safe_key = validate_safe_name(key, label="Ação de Registro")
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

        mode = str(item.get("mode") or "set").casefold()
        safe_path = quote_powershell_literal(path)
        safe_name = quote_powershell_literal(name)

        if mode == "remove":
            script = f"""
if (Test-Path {safe_path}) {{
    Remove-ItemProperty -Path {safe_path} -Name {safe_name} -ErrorAction Stop
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
            value = item.get("value")
            if isinstance(value, str):
                ps_value = quote_powershell_literal(value)
            elif isinstance(value, bool):
                ps_value = "$true" if value else "$false"
            elif isinstance(value, (int, float)):
                ps_value = str(value)
            elif isinstance(value, list):
                items = ",".join(
                    quote_powershell_literal(str(entry))
                    for entry in value
                )
                ps_value = f"@({items})"
            else:
                return CommandResult.failure(
                    host,
                    safe_key,
                    "Tipo de valor não suportado.",
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
