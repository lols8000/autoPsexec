from __future__ import annotations

from core.jobs import OperationClass
from ..models import ExecutionAction, ExecutionParameter, ParameterKind, RiskLevel
from .common import ExecutionDependencies, _process_absent, _register


def register(registry, deps: ExecutionDependencies) -> None:
    _register(
        registry,
        ExecutionAction(
            "process.kill_name",
            "Finalizar processo por nome",
            "processes",
            "Processos",
            "Finaliza todas as instâncias do processo informado.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.MEDIUM,
            "Pode fechar aplicação e causar perda de trabalho não salvo.",
            120,
            destructive=True,
            parameters=(
                ExecutionParameter("process_name", "Nome do processo", ParameterKind.TEXT),
            ),
        ),
        lambda host, p: deps.system.kill_process(host, p["process_name"]),
        before_probe=lambda host, p: deps.system.process_status(
            host,
            process_name=p["process_name"],
        ),
        after_probe=lambda host, p: deps.system.process_status(
            host,
            process_name=p["process_name"],
        ),
        validator=_process_absent,
    )
    _register(
        registry,
        ExecutionAction(
            "process.kill_pid",
            "Finalizar processo por PID",
            "processes",
            "Processos",
            "Finaliza uma instância específica pelo PID.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.MEDIUM,
            "Pode encerrar aplicação ou componente do sistema.",
            120,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "pid",
                    "PID",
                    ParameterKind.INTEGER,
                    min_value=1,
                    max_value=2_147_483_647,
                ),
            ),
        ),
        lambda host, p: deps.system.kill_process_pid(host, p["pid"]),
        before_probe=lambda host, p: deps.system.process_status(
            host,
            pid=p["pid"],
        ),
        after_probe=lambda host, p: deps.system.process_status(
            host,
            pid=p["pid"],
        ),
        validator=_process_absent,
    )
