from __future__ import annotations

from core.jobs import OperationClass

from ..models import ExecutionAction, RiskLevel
from .common import ExecutionDependencies, _defender_ready, _register


def register(registry, deps: ExecutionDependencies) -> None:
    _register(
        registry,
        ExecutionAction(
            "defender.signatures",
            "Atualizar assinaturas do Defender",
            "security",
            "Segurança / Defender",
            "Executa Update-MpSignature.",
            OperationClass.LIGHT_WRITE,
            RiskLevel.LOW,
            "Baixo impacto; requer Defender disponível.",
            900,
            required_capabilities=("Defender",),
            tags=("defender", "assinatura", "signature", "update"),
        ),
        lambda host, p: deps.security.update_defender_signatures(host),
        after_probe=lambda host, p: deps.security.defender_status(host),
        validator=_defender_ready,
    )
    for key, title, scan_type, timeout, risk in (
        (
            "defender.quick_scan",
            "Executar verificação rápida",
            "QuickScan",
            2400,
            RiskLevel.LOW,
        ),
        (
            "defender.full_scan",
            "Executar verificação completa",
            "FullScan",
            7800,
            RiskLevel.MEDIUM,
        ),
    ):
        _register(
            registry,
            ExecutionAction(
                key,
                title,
                "security",
                "Segurança / Defender",
                title,
                OperationClass.HEAVY_WRITE,
                risk,
                "Pode consumir CPU/disco durante a varredura.",
                timeout,
                required_capabilities=("Defender",),
                tags=("defender", "scan", scan_type.casefold()),
            ),
            lambda host, p, st=scan_type: deps.security.defender_scan(
                host,
                st,
            ),
            before_probe=lambda host, p: deps.security.defender_status(host),
            after_probe=lambda host, p: deps.security.defender_status(host),
            validator=_defender_ready,
        )
