from __future__ import annotations

from core.jobs import OperationClass
from ..models import ExecutionAction, ExecutionParameter, ParameterKind, RiskLevel
from ..validators import command_field_true
from .common import ExecutionDependencies, _register


def register(registry, deps: ExecutionDependencies) -> None:
    certificate_keys = deps.certificates.keys()
    if certificate_keys:
        _register(
            registry,
            ExecutionAction(
                "certificate.import",
                "Importar certificado homologado",
                "certificates",
                "Certificados",
                "Importa .cer/.crt permitido em LocalMachine.",
                OperationClass.HEAVY_WRITE,
                RiskLevel.HIGH,
                "Altera stores de confiança do computador.",
                300,
                parameters=(
                    ExecutionParameter(
                        "certificate_key",
                        "Certificado",
                        ParameterKind.CHOICE,
                        choices=certificate_keys,
                    ),
                ),
            ),
            lambda host, p: deps.certificates.import_certificate(
                host,
                p["certificate_key"],
            ),
            validator=command_field_true(
                "Imported",
                pass_message="Certificado importado.",
                fail_message="Importação não foi confirmada.",
            ),
        )
