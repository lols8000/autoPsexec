from __future__ import annotations

from core.jobs import OperationClass
from remediation import validate_spooler

from ..models import ExecutionAction, ExecutionParameter, ParameterKind, RiskLevel
from ..validators import command_completed
from .common import ExecutionDependencies, _printer_exists, _register, _wrap_three_arg


def register(registry, deps: ExecutionDependencies) -> None:
    _register(
        registry,
        ExecutionAction(
            "printer.restart_spooler",
            "Reiniciar Spooler",
            "printers",
            "Impressão",
            "Reinicia Spooler e valida Running.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Jobs em processamento podem ser interrompidos temporariamente.",
            180,
        ),
        lambda host, p: deps.printers.restart_spooler(host),
        before_probe=lambda host, p: deps.printers.spooler_status(host),
        after_probe=lambda host, p: deps.printers.spooler_status(host),
        validator=_wrap_three_arg(validate_spooler),
    )
    _register(
        registry,
        ExecutionAction(
            "printer.clear_queue",
            "Limpar fila de impressão",
            "printers",
            "Impressão",
            "Remove jobs de uma impressora ou de todas quando o nome ficar vazio.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Trabalhos pendentes serão perdidos.",
            180,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "printer_name",
                    "Impressora (vazio = todas)",
                    ParameterKind.TEXT,
                    required=False,
                ),
            ),
        ),
        lambda host, p: deps.printers.clear_queue(
            host,
            p.get("printer_name"),
        ),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "printer.add_connection",
            "Adicionar impressora compartilhada",
            "printers",
            "Impressão",
            r"Adiciona conexão \\\\servidor\\fila.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.MEDIUM,
            "Adiciona impressora ao sistema.",
            240,
            parameters=(
                ExecutionParameter(
                    "connection_name",
                    "UNC da impressora",
                    ParameterKind.TEXT,
                ),
            ),
        ),
        lambda host, p: deps.printers.add_connection(
            host,
            p["connection_name"],
        ),
        after_probe=lambda host, p: deps.printers.printer_status(
            host,
            p["connection_name"],
        ),
        validator=_printer_exists(True),
    )
    _register(
        registry,
        ExecutionAction(
            "printer.remove",
            "Remover impressora",
            "printers",
            "Impressão",
            "Remove uma impressora pelo nome exato.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.HIGH,
            "Remove configuração de impressão da estação.",
            180,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "printer_name",
                    "Nome da impressora",
                    ParameterKind.TEXT,
                ),
            ),
        ),
        lambda host, p: deps.printers.remove_printer(
            host,
            p["printer_name"],
        ),
        before_probe=lambda host, p: deps.printers.printer_status(
            host,
            p["printer_name"],
        ),
        after_probe=lambda host, p: deps.printers.printer_status(
            host,
            p["printer_name"],
        ),
        validator=_printer_exists(False),
    )
