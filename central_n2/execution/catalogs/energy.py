from __future__ import annotations

from core.jobs import OperationClass
from ..models import ExecutionAction, ExecutionParameter, ParameterKind, RiskLevel
from ..validators import command_completed
from .common import ExecutionDependencies, _register


def register(registry, deps: ExecutionDependencies) -> None:
    _register(
        registry,
        ExecutionAction(
            "energy.restart",
            "Reiniciar estação",
            "energy",
            "Energia / Sessões",
            "Agenda reinicialização administrativa.",
            OperationClass.DISRUPTIVE,
            RiskLevel.CRITICAL,
            "Interrompe sessões e indisponibiliza a estação temporariamente.",
            180,
            may_break_connectivity=True,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "delay_seconds",
                    "Atraso em segundos",
                    ParameterKind.INTEGER,
                    required=False,
                    default=0,
                    min_value=0,
                    max_value=3600,
                ),
            ),
        ),
        lambda host, p: deps.system.restart(
            host,
            p["delay_seconds"],
        ),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "energy.shutdown",
            "Desligar estação",
            "energy",
            "Energia / Sessões",
            "Agenda desligamento administrativo.",
            OperationClass.DISRUPTIVE,
            RiskLevel.CRITICAL,
            "Interrompe sessões e desliga a estação.",
            180,
            may_break_connectivity=True,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "delay_seconds",
                    "Atraso em segundos",
                    ParameterKind.INTEGER,
                    required=False,
                    default=0,
                    min_value=0,
                    max_value=3600,
                ),
            ),
        ),
        lambda host, p: deps.system.shutdown(
            host,
            p["delay_seconds"],
        ),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "energy.abort",
            "Cancelar shutdown/restart pendente",
            "energy",
            "Energia / Sessões",
            "Executa shutdown /a.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Cancela temporizador de energia pendente.",
            120,
        ),
        lambda host, p: deps.system.abort_shutdown(host),
        validator=command_completed,
    )
    _register(
        registry,
        ExecutionAction(
            "session.message",
            "Enviar mensagem para usuários",
            "energy",
            "Energia / Sessões",
            "Envia mensagem para sessões interativas.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Exibe mensagem aos usuários conectados.",
            120,
            parameters=(
                ExecutionParameter(
                    "message",
                    "Mensagem",
                    ParameterKind.TEXT,
                ),
            ),
        ),
        lambda host, p: deps.system.send_message(
            host,
            p["message"],
        ),
        validator=command_completed,
    )
