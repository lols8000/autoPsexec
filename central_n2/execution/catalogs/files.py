from __future__ import annotations

from typing import Any

from core.jobs import OperationClass
from core.result import CommandResult
from remediation import ValidationResult, ValidationStatus

from ..models import (
    ExecutionAction,
    ExecutionParameter,
    ParameterKind,
    RiskLevel,
)
from ..policy import PolicyCheck, PolicyState
from ..validators import command_field_true
from .common import ExecutionDependencies, _register


def _payload(value: Any) -> dict[str, Any]:
    if isinstance(value, CommandResult):
        return value.data if isinstance(value.data, dict) else {}
    return value if isinstance(value, dict) else {}


def register(registry, deps: ExecutionDependencies) -> None:
    _register(
        registry,
        ExecutionAction(
            "file.ensure_directory",
            "Criar / garantir diretório",
            "files",
            "Arquivos",
            "Cria diretório remoto validado.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Cria estrutura de diretórios.",
            180,
            parameters=(
                ExecutionParameter("path", "Caminho", ParameterKind.TEXT),
            ),
            idempotent=True,
            tags=("arquivo", "diretório", "folder", "mkdir"),
        ),
        lambda host, p: deps.files.ensure_directory(
            host,
            p["path"],
        ),
        validator=command_field_true(
            "Exists",
            pass_message="Diretório confirmado.",
            fail_message="Diretório não foi confirmado.",
        ),
    )

    def destination_must_be_absent(context, parameters):
        result = deps.files.path_status(
            context.host,
            parameters["source"],
            parameters["destination"],
        )
        if not result.success or not isinstance(result.data, dict):
            return [
                PolicyCheck(
                    "file.move.destination",
                    PolicyState.FAIL,
                    "Não foi possível validar origem/destino antes do move.",
                    result.stderr,
                )
            ]
        if result.data.get("SourceExists") is not True:
            return [
                PolicyCheck(
                    "file.move.source",
                    PolicyState.FAIL,
                    "A origem não existe.",
                    result.data,
                )
            ]
        if result.data.get("DestinationExists") is True:
            return [
                PolicyCheck(
                    "file.move.destination",
                    PolicyState.FAIL,
                    "O destino já existe; move bloqueado para evitar overwrite/merge ambíguo.",
                    result.data,
                )
            ]
        return [
            PolicyCheck(
                "file.move.destination",
                PolicyState.PASS,
                "Origem existe e destino está livre.",
                result.data,
            )
        ]

    def validate_move(before, command, after, parameters):
        data = _payload(after)
        if command.indeterminate:
            return ValidationResult(
                ValidationStatus.UNKNOWN,
                "Move teve resultado indeterminado.",
                data,
            )
        if (
            data.get("SourceExists") is False
            and data.get("DestinationExists") is True
        ):
            return ValidationResult(
                ValidationStatus.PASS,
                "Movimentação confirmada no destino.",
                data,
            )
        return ValidationResult(
            ValidationStatus.FAIL if not command.success else ValidationStatus.UNKNOWN,
            "Estado final do move não foi confirmado.",
            data,
        )

    def rollback_move(host, parameters, before):
        return deps.files.move_path(
            host,
            parameters["destination"],
            parameters["source"],
        )

    def rollback_path_must_be_safe(context, parameters):
        result = deps.files.path_status(
            context.host,
            parameters["source"],
            parameters["destination"],
        )
        if not result.success or not isinstance(result.data, dict):
            return [
                PolicyCheck(
                    "file.move.rollback",
                    PolicyState.FAIL,
                    "Não foi possível validar os caminhos antes do rollback.",
                    result.stderr,
                )
            ]
        if result.data.get("DestinationExists") is not True:
            return [
                PolicyCheck(
                    "file.move.rollback.destination",
                    PolicyState.FAIL,
                    "O destino atual não existe; não há o que mover de volta.",
                    result.data,
                )
            ]
        if result.data.get("SourceExists") is True:
            return [
                PolicyCheck(
                    "file.move.rollback.source",
                    PolicyState.FAIL,
                    "A origem original voltou a existir; rollback bloqueado para evitar overwrite.",
                    result.data,
                )
            ]
        return [
            PolicyCheck(
                "file.move.rollback",
                PolicyState.PASS,
                "Destino existe e origem original está livre para rollback.",
                result.data,
            )
        ]

    def validate_move_rollback(before, command, after, parameters):
        data = _payload(after)
        if command.indeterminate:
            return ValidationResult(
                ValidationStatus.UNKNOWN,
                "Rollback do move teve resultado indeterminado.",
                data,
            )
        if (
            data.get("SourceExists") is True
            and data.get("DestinationExists") is False
        ):
            return ValidationResult(
                ValidationStatus.PASS,
                "Caminho original restaurado.",
                data,
            )
        return ValidationResult(
            ValidationStatus.FAIL if not command.success else ValidationStatus.UNKNOWN,
            "Caminho original não foi confirmado.",
            data,
        )

    _register(
        registry,
        ExecutionAction(
            "file.move",
            "Mover / renomear caminho",
            "files",
            "Arquivos",
            "Move arquivo ou diretório apenas quando o destino ainda não existe.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Muda localização de dados e pode afetar aplicações.",
            600,
            destructive=True,
            parameters=(
                ExecutionParameter("source", "Origem", ParameterKind.TEXT),
                ExecutionParameter("destination", "Destino", ParameterKind.TEXT),
            ),
            rollback_strategy="Mover o destino de volta para a origem original.",
            tags=("arquivo", "mover", "renomear", "move"),
        ),
        lambda host, p: deps.files.move_path(
            host,
            p["source"],
            p["destination"],
        ),
        before_probe=lambda host, p: deps.files.path_status(
            host,
            p["source"],
            p["destination"],
        ),
        after_probe=lambda host, p: deps.files.path_status(
            host,
            p["source"],
            p["destination"],
        ),
        validator=validate_move,
        preconditions=(destination_must_be_absent,),
        rollback_preconditions=(rollback_path_must_be_safe,),
        rollback_handler=rollback_move,
        rollback_validator=validate_move_rollback,
    )

    _register(
        registry,
        ExecutionAction(
            "file.remove",
            "Excluir arquivo",
            "files",
            "Arquivos",
            "Exclui somente arquivo; não remove diretório recursivamente.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.CRITICAL,
            "Exclusão permanente de arquivo.",
            300,
            destructive=True,
            parameters=(
                ExecutionParameter("path", "Arquivo", ParameterKind.TEXT),
            ),
            tags=("arquivo", "excluir", "delete", "remove"),
        ),
        lambda host, p: deps.files.remove_file(
            host,
            p["path"],
        ),
        validator=command_field_true(
            "Removed",
            pass_message="Arquivo removido e ausência confirmada.",
            fail_message="Remoção não foi confirmada.",
        ),
    )
