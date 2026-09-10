from __future__ import annotations

from typing import Any

from core.result import CommandResult
from .engine import ValidationResult, ValidationStatus


def _result_data(value: Any) -> dict[str, Any]:
    if isinstance(value, CommandResult):
        return value.data if isinstance(value.data, dict) else {}
    return value if isinstance(value, dict) else {}


def validate_spooler(
    before: Any,
    command: CommandResult,
    after: Any,
) -> ValidationResult:
    data = _result_data(after)
    status = str(data.get("Status") or "").casefold()

    if status == "running":
        return ValidationResult(
            ValidationStatus.PASS,
            "Spooler está em execução após a remediação.",
            data,
        )
    if command.indeterminate:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "Não foi possível confirmar o estado final do Spooler.",
            data,
        )
    return ValidationResult(
        ValidationStatus.FAIL,
        "Spooler não foi confirmado como Running.",
        data,
    )


def validate_cleanup(
    before: Any,
    command: CommandResult,
    after: Any,
) -> ValidationResult:
    data = _result_data(command)
    recovered = data.get("RecoveredGB")
    if command.success and recovered is not None:
        return ValidationResult(
            ValidationStatus.PASS,
            f"Limpeza concluída; {recovered} GB recuperado(s).",
            data,
        )
    if command.indeterminate:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "A limpeza pode ter sido executada, mas o resultado não foi confirmado.",
            data,
        )
    return ValidationResult(
        ValidationStatus.FAIL,
        "A limpeza não foi confirmada.",
        data,
    )


def validate_windows_update_reset(
    before: Any,
    command: CommandResult,
    after: Any,
) -> ValidationResult:
    data = _result_data(command)
    restored = data.get("RestoredOriginalRunningServices")
    if command.success and restored is True:
        return ValidationResult(
            ValidationStatus.PASS,
            "Componentes resetados e serviços originalmente ativos restaurados.",
            data,
        )
    if command.indeterminate or restored is None:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "Não foi possível confirmar integralmente o estado dos serviços do Windows Update.",
            data,
        )
    return ValidationResult(
        ValidationStatus.FAIL,
        "Serviços do Windows Update não retornaram ao estado esperado.",
        data,
    )


def validate_gpupdate(
    before: Any,
    command: CommandResult,
    after: Any,
) -> ValidationResult:
    if not command.success:
        return ValidationResult(
            ValidationStatus.UNKNOWN if command.indeterminate else ValidationStatus.FAIL,
            (
                "GPUpdate teve resultado indeterminado."
                if command.indeterminate
                else "GPUpdate falhou."
            ),
            after,
        )

    if isinstance(after, CommandResult) and after.success:
        return ValidationResult(
            ValidationStatus.PASS,
            "GPUpdate concluiu e GPResult foi coletado após a aplicação.",
            {"gpresult": after.stdout},
        )

    if isinstance(after, dict) and after.get("_probe_success") is True:
        return ValidationResult(
            ValidationStatus.PASS,
            "GPUpdate concluiu e GPResult foi coletado após a aplicação.",
            {"gpresult": after.get("_stdout", "")},
        )

    return ValidationResult(
        ValidationStatus.UNKNOWN,
        "GPUpdate concluiu, mas a validação por GPResult não foi obtida.",
        after,
    )
