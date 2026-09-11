from __future__ import annotations

from core.jobs import OperationClass

from ..models import ExecutionAction, ExecutionParameter, ParameterKind, RiskLevel
from ..validators import command_field_true
from .common import ExecutionDependencies, _register


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
    _register(
        registry,
        ExecutionAction(
            "file.move",
            "Mover / renomear caminho",
            "files",
            "Arquivos",
            "Move arquivo ou diretório entre caminhos absolutos.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Muda localização de dados e pode afetar aplicações.",
            600,
            destructive=True,
            parameters=(
                ExecutionParameter("source", "Origem", ParameterKind.TEXT),
                ExecutionParameter("destination", "Destino", ParameterKind.TEXT),
            ),
        ),
        lambda host, p: deps.files.move_path(
            host,
            p["source"],
            p["destination"],
        ),
        validator=command_field_true(
            "DestinationExists",
            pass_message="Destino confirmado após a movimentação.",
            fail_message="Destino não foi confirmado.",
        ),
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
