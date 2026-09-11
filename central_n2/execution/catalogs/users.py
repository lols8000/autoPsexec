from __future__ import annotations

from core.jobs import OperationClass
from remediation import validate_cleanup

from ..models import (
    ExecutionAction,
    ExecutionParameter,
    ParameterKind,
    RiskLevel,
    SelectorKind,
)
from ..validators import command_completed
from .common import (
    ExecutionDependencies,
    _profile_removed,
    _register,
    _wrap_three_arg,
)


def register(registry, deps: ExecutionDependencies) -> None:
    _register(
        registry,
        ExecutionAction(
            "session.logoff",
            "Encerrar sessão de usuário",
            "users",
            "Usuários / Perfis",
            "Executa logoff pelo ID da sessão.",
            OperationClass.DISRUPTIVE,
            RiskLevel.HIGH,
            "Aplicações abertas podem perder trabalho não salvo.",
            180,
            destructive=True,
            parameters=(
                ExecutionParameter(
                    "session_id",
                    "Sessão",
                    ParameterKind.INTEGER,
                    min_value=0,
                    max_value=65535,
                    selector=SelectorKind.SESSION,
                ),
            ),
            tags=("usuário", "sessão", "logoff"),
        ),
        lambda host, p: deps.system.logoff_session(
            host,
            p["session_id"],
        ),
        validator=command_completed,
    )
    profile_parameter = ExecutionParameter(
        "sid",
        "Perfil",
        ParameterKind.TEXT,
        selector=SelectorKind.PROFILE,
    )
    _register(
        registry,
        ExecutionAction(
            "profile.clean_temp",
            "Limpar TEMP de perfil",
            "users",
            "Usuários / Perfis",
            "Limpa AppData\\Local\\Temp do SID selecionado.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.MEDIUM,
            "Arquivos temporários em uso são preservados quando bloqueados.",
            600,
            parameters=(profile_parameter,),
            tags=("usuário", "perfil", "temp", "limpeza"),
        ),
        lambda host, p: deps.users.clean_profile_temp(
            host,
            p["sid"],
        ),
        before_probe=lambda host, p: deps.users.profile_status(
            host,
            p["sid"],
        ),
        after_probe=lambda host, p: deps.users.profile_status(
            host,
            p["sid"],
        ),
        validator=_wrap_three_arg(validate_cleanup),
    )
    _register(
        registry,
        ExecutionAction(
            "profile.remove",
            "Remover perfil de usuário",
            "users",
            "Usuários / Perfis",
            "Remove somente perfil não carregado e não especial.",
            OperationClass.HEAVY_WRITE,
            RiskLevel.CRITICAL,
            "Dados locais do perfil podem ser removidos permanentemente.",
            600,
            destructive=True,
            parameters=(profile_parameter,),
            tags=("usuário", "perfil", "remover"),
        ),
        lambda host, p: deps.users.remove_profile(host, p["sid"]),
        before_probe=lambda host, p: deps.users.profile_status(
            host,
            p["sid"],
        ),
        after_probe=lambda host, p: deps.users.profile_status(
            host,
            p["sid"],
        ),
        validator=_profile_removed,
    )
