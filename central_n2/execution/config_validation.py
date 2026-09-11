from __future__ import annotations

import copy
from dataclasses import dataclass, field
from pathlib import PureWindowsPath
from typing import Any


@dataclass(frozen=True, slots=True)
class CatalogIssue:
    section: str
    key: str
    message: str


@dataclass(slots=True)
class CatalogValidationReport:
    settings: dict[str, Any]
    issues: list[CatalogIssue] = field(default_factory=list)
    enabled_counts: dict[str, int] = field(default_factory=dict)

    @property
    def valid(self) -> bool:
        return not self.issues


_ALLOWED_PACKAGE_TYPES = {"msi", "exe", "cmd", "bat"}
_ALLOWED_CERTIFICATE_STORES = {"Root", "CA", "My", "TrustedPeople"}
_ALLOWED_REGISTRY_TYPES = {
    "String",
    "ExpandString",
    "DWord",
    "QWord",
    "MultiString",
    "Binary",
}


def _absolute_windows_path(value: str) -> bool:
    try:
        return bool(PureWindowsPath(value).drive)
    except (TypeError, ValueError):
        return False


def _validate_packages(
    source: Any,
    issues: list[CatalogIssue],
) -> dict[str, Any]:
    if source in (None, {}):
        return {}
    if not isinstance(source, dict):
        issues.append(
            CatalogIssue(
                "packages",
                "*",
                "packages deve ser um objeto chaveado por identificador.",
            )
        )
        return {}

    valid: dict[str, Any] = {}
    for key, item in source.items():
        if not isinstance(item, dict):
            issues.append(
                CatalogIssue("packages", str(key), "Entrada deve ser objeto.")
            )
            continue

        path = str(item.get("source") or "").strip()
        package_type = str(item.get("type") or "").casefold()
        timeout = item.get("timeout_seconds", 1800)

        errors: list[str] = []
        if not path:
            errors.append("source obrigatório")
        if package_type not in _ALLOWED_PACKAGE_TYPES:
            errors.append(
                "type deve ser msi, exe, cmd ou bat"
            )
        try:
            if int(timeout) <= 0:
                errors.append("timeout_seconds deve ser > 0")
        except (TypeError, ValueError):
            errors.append("timeout_seconds deve ser inteiro")

        remote_path = str(item.get("remote_path") or "").strip()
        if remote_path and not _absolute_windows_path(remote_path):
            errors.append("remote_path deve ser absoluto")

        if errors:
            issues.append(
                CatalogIssue(
                    "packages",
                    str(key),
                    "; ".join(errors),
                )
            )
            continue

        valid[str(key)] = dict(item)
    return valid


def _validate_certificates(
    source: Any,
    issues: list[CatalogIssue],
) -> dict[str, Any]:
    if source in (None, {}):
        return {}
    if not isinstance(source, dict):
        issues.append(
            CatalogIssue(
                "certificates",
                "*",
                "certificates deve ser um objeto chaveado por identificador.",
            )
        )
        return {}

    valid: dict[str, Any] = {}
    for key, item in source.items():
        if not isinstance(item, dict):
            issues.append(
                CatalogIssue(
                    "certificates",
                    str(key),
                    "Entrada deve ser objeto.",
                )
            )
            continue

        path = str(item.get("source") or "").strip()
        store = str(item.get("store") or "Root")
        errors: list[str] = []

        if not path.casefold().endswith((".cer", ".crt")):
            errors.append("source deve terminar em .cer ou .crt")
        if store not in _ALLOWED_CERTIFICATE_STORES:
            errors.append(
                "store permitido: Root, CA, My ou TrustedPeople"
            )

        remote_path = str(item.get("remote_path") or "").strip()
        if remote_path and not _absolute_windows_path(remote_path):
            errors.append("remote_path deve ser absoluto")

        if errors:
            issues.append(
                CatalogIssue(
                    "certificates",
                    str(key),
                    "; ".join(errors),
                )
            )
            continue

        valid[str(key)] = dict(item)
    return valid


def _validate_registry_actions(
    source: Any,
    issues: list[CatalogIssue],
) -> dict[str, Any]:
    if source in (None, {}):
        return {}
    if not isinstance(source, dict):
        issues.append(
            CatalogIssue(
                "registry_actions",
                "*",
                "registry_actions deve ser um objeto chaveado por identificador.",
            )
        )
        return {}

    valid: dict[str, Any] = {}
    for key, item in source.items():
        if not isinstance(item, dict):
            issues.append(
                CatalogIssue(
                    "registry_actions",
                    str(key),
                    "Entrada deve ser objeto.",
                )
            )
            continue

        path = str(item.get("path") or "").strip()
        name = str(item.get("name") or "").strip()
        mode = str(item.get("mode") or "set").casefold()
        value_type = str(item.get("type") or "String")
        errors: list[str] = []

        if not path.upper().startswith("HKLM:\\"):
            errors.append("path deve permanecer em HKLM")
        if not name:
            errors.append("name obrigatório")
        if mode not in {"set", "remove"}:
            errors.append("mode deve ser set ou remove")
        if mode == "set":
            if value_type not in _ALLOWED_REGISTRY_TYPES:
                errors.append("type de Registro não permitido")
            if "value" not in item:
                errors.append("value obrigatório para mode=set")

        if errors:
            issues.append(
                CatalogIssue(
                    "registry_actions",
                    str(key),
                    "; ".join(errors),
                )
            )
            continue

        valid[str(key)] = dict(item)
    return valid


def validate_execution_configuration(
    settings: dict[str, Any],
) -> CatalogValidationReport:
    sanitized = copy.deepcopy(settings)
    issues: list[CatalogIssue] = []

    sanitized["packages"] = _validate_packages(
        sanitized.get("packages"),
        issues,
    )
    sanitized["certificates"] = _validate_certificates(
        sanitized.get("certificates"),
        issues,
    )
    sanitized["registry_actions"] = _validate_registry_actions(
        sanitized.get("registry_actions"),
        issues,
    )

    counts = {
        section: len(sanitized.get(section, {}))
        for section in (
            "packages",
            "certificates",
            "registry_actions",
        )
    }

    sanitized["_execution_catalog_validation"] = {
        "issues": [
            {
                "section": issue.section,
                "key": issue.key,
                "message": issue.message,
            }
            for issue in issues
        ],
        "enabled_counts": counts,
    }

    return CatalogValidationReport(
        settings=sanitized,
        issues=issues,
        enabled_counts=counts,
    )
