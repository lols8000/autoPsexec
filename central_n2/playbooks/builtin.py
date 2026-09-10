from __future__ import annotations

from .base import PlaybookSpec, PlaybookStep


def _step(key: str, label: str) -> PlaybookStep:
    return PlaybookStep(key, label)


def builtin_playbooks() -> dict[str, PlaybookSpec]:
    specs = (
        PlaybookSpec(
            "slow",
            "Computador lento",
            (
                _step("health", "Saúde geral"),
                _step("performance", "Performance"),
                _step("disk", "Disco"),
                _step("processes", "Processos"),
                _step("startup", "Inicialização"),
            ),
        ),
        PlaybookSpec(
            "network",
            "Sem internet / rede",
            (
                _step("network", "IP/DNS/Gateway"),
                _step("adapters", "Adaptadores"),
                _step("connections", "Conexões TCP"),
                _step("proxy", "Proxy"),
            ),
        ),
        PlaybookSpec(
            "printer",
            "Não imprime",
            (
                _step("printers", "Impressoras"),
                _step("print_queue", "Fila"),
                _step("services", "Serviços"),
            ),
        ),
        PlaybookSpec(
            "domain",
            "Problema de domínio / GPO",
            (
                _step("domain", "Domínio"),
                _step("gpresult", "GPResult"),
            ),
        ),
        PlaybookSpec(
            "update",
            "Windows Update",
            (
                _step("updates", "Windows Update"),
                _step("health", "Saúde"),
            ),
        ),
        PlaybookSpec(
            "crash",
            "Aplicativo fechando",
            (
                _step("app_crashes", "Crashes"),
                _step("processes", "Processos"),
            ),
        ),
        PlaybookSpec(
            "bsod",
            "Tela azul / BSOD",
            (
                _step("bsod", "BugChecks/dumps"),
                _step("devices", "Dispositivos"),
            ),
        ),
        PlaybookSpec(
            "disk",
            "Disco cheio",
            (
                _step("disk", "Espaço"),
                _step("profiles", "Perfis"),
                _step("cleanup_estimate", "Estimativa limpeza"),
            ),
        ),
        PlaybookSpec(
            "glpi",
            "GLPI Agent",
            (
                _step("glpi", "Status"),
                _step("glpi_log", "Log recente"),
            ),
        ),
    )
    return {spec.key: spec for spec in specs}
