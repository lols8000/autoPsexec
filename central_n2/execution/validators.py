from __future__ import annotations

from typing import Any, Callable

from core.result import CommandResult
from remediation import ValidationResult, ValidationStatus


def command_completed(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    if command.indeterminate:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "A ação pode ter sido entregue, mas o resultado final não foi confirmado.",
            after,
        )
    if command.success:
        return ValidationResult(
            ValidationStatus.PASS,
            "A execução foi concluída sem erro de transporte/comando.",
            after,
        )
    return ValidationResult(
        ValidationStatus.FAIL,
        command.stderr or "A execução falhou.",
        after,
    )


def field_equals(
    field: str,
    expected: Any,
    *,
    pass_message: str,
    fail_message: str,
) -> Callable[
    [Any, CommandResult, Any, dict[str, Any]],
    ValidationResult,
]:
    def validate(
        before: Any,
        command: CommandResult,
        after: Any,
        parameters: dict[str, Any],
    ) -> ValidationResult:
        if command.indeterminate:
            return ValidationResult(
                ValidationStatus.UNKNOWN,
                "A execução teve resultado indeterminado.",
                after,
            )

        data = after
        if isinstance(after, CommandResult):
            data = after.data

        if not isinstance(data, dict) or field not in data:
            if not command.success:
                return ValidationResult(
                    ValidationStatus.FAIL,
                    command.stderr or fail_message,
                    data,
                )
            return ValidationResult(
                ValidationStatus.UNKNOWN,
                f"Não foi possível validar o campo '{field}'.",
                data,
            )

        if data.get(field) == expected:
            return ValidationResult(
                ValidationStatus.PASS,
                pass_message,
                data,
            )

        return ValidationResult(
            ValidationStatus.FAIL,
            fail_message,
            data,
        )

    return validate


def service_running(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    data = after.data if isinstance(after, CommandResult) else after
    if command.indeterminate:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "O serviço pode ter sido alterado, mas o estado final não foi confirmado.",
            data,
        )
    if isinstance(data, dict):
        status = str(data.get("Status") or "").casefold()
        if status == "running":
            return ValidationResult(
                ValidationStatus.PASS,
                "Serviço confirmado como Running.",
                data,
            )
        return ValidationResult(
            ValidationStatus.FAIL,
            "Serviço não foi confirmado como Running.",
            data,
        )
    return command_completed(before, command, after, parameters)


def service_stopped(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    data = after.data if isinstance(after, CommandResult) else after
    if command.indeterminate:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "O serviço pode ter sido alterado, mas o estado final não foi confirmado.",
            data,
        )
    if isinstance(data, dict):
        status = str(data.get("Status") or "").casefold()
        if status == "stopped":
            return ValidationResult(
                ValidationStatus.PASS,
                "Serviço confirmado como Stopped.",
                data,
            )
        return ValidationResult(
            ValidationStatus.FAIL,
            "Serviço não foi confirmado como Stopped.",
            data,
        )
    return command_completed(before, command, after, parameters)



def command_field_true(
    field: str,
    *,
    pass_message: str,
    fail_message: str,
):
    def validate(
        before: Any,
        command: CommandResult,
        after: Any,
        parameters: dict[str, Any],
    ) -> ValidationResult:
        if command.indeterminate:
            return ValidationResult(
                ValidationStatus.UNKNOWN,
                "A execução teve resultado indeterminado.",
                command.data,
            )
        data = command.data
        if not isinstance(data, dict) or field not in data:
            if command.success:
                return ValidationResult(
                    ValidationStatus.UNKNOWN,
                    f"Não foi possível validar o campo '{field}'.",
                    data,
                )
            return ValidationResult(
                ValidationStatus.FAIL,
                command.stderr or fail_message,
                data,
            )
        if bool(data.get(field)):
            return ValidationResult(
                ValidationStatus.PASS,
                pass_message,
                data,
            )
        return ValidationResult(
            ValidationStatus.FAIL,
            fail_message,
            data,
        )

    return validate


def after_field_matches_parameter(
    field: str,
    parameter_key: str,
    *,
    pass_message: str,
    fail_message: str,
):
    def validate(
        before: Any,
        command: CommandResult,
        after: Any,
        parameters: dict[str, Any],
    ) -> ValidationResult:
        if command.indeterminate:
            return ValidationResult(
                ValidationStatus.UNKNOWN,
                "A execução teve resultado indeterminado.",
                after,
            )
        data = after.data if isinstance(after, CommandResult) else after
        if not isinstance(data, dict) or field not in data:
            return ValidationResult(
                ValidationStatus.UNKNOWN if command.success else ValidationStatus.FAIL,
                (
                    f"Não foi possível validar o campo '{field}'."
                    if command.success
                    else command.stderr or fail_message
                ),
                data,
            )
        expected = parameters.get(parameter_key)
        if str(data.get(field)).casefold() == str(expected).casefold():
            return ValidationResult(
                ValidationStatus.PASS,
                pass_message,
                data,
            )
        return ValidationResult(
            ValidationStatus.FAIL,
            fail_message,
            data,
        )

    return validate



def postcheck_succeeded(
    before: Any,
    command: CommandResult,
    after: Any,
    parameters: dict[str, Any],
) -> ValidationResult:
    if command.indeterminate:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "A execução teve resultado indeterminado.",
            after,
        )

    if isinstance(after, CommandResult):
        if after.success:
            return ValidationResult(
                ValidationStatus.PASS,
                "Postcheck executado com sucesso.",
                after.data if after.data is not None else after.stdout,
            )
        return ValidationResult(
            ValidationStatus.UNKNOWN if command.success else ValidationStatus.FAIL,
            after.stderr or "Postcheck falhou.",
            after.data,
        )

    if isinstance(after, dict):
        probe_success = after.get("_probe_success")
        if probe_success is True:
            return ValidationResult(
                ValidationStatus.PASS,
                "Postcheck executado com sucesso.",
                after,
            )
        if probe_success is False:
            return ValidationResult(
                ValidationStatus.UNKNOWN if command.success else ValidationStatus.FAIL,
                str(after.get("_error") or "Postcheck não confirmou o estado."),
                after,
            )

    if command.success:
        return ValidationResult(
            ValidationStatus.UNKNOWN,
            "Comando concluiu, mas não houve evidência de postcheck suficiente.",
            after,
        )

    return ValidationResult(
        ValidationStatus.FAIL,
        command.stderr or "Execução falhou.",
        after,
    )
