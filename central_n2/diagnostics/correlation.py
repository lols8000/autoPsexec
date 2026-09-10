from __future__ import annotations

from .models import Diagnosis, Finding


class CorrelationEngine:
    def correlate(self, findings: list[Finding]) -> list[Diagnosis]:
        by = {finding.id: finding for finding in findings}
        diagnoses: list[Diagnosis] = []

        storage = [
            by[key]
            for key in ("DISK_CRITICAL", "DISK_LOW", "DISK_SATURATION")
            if key in by
        ]
        if "DISK_CRITICAL" in by and "DISK_SATURATION" in by:
            diagnoses.append(
                Diagnosis(
                    "STORAGE_PRESSURE",
                    "Pressão de armazenamento",
                    "alta",
                    "Espaço livre crítico combinado com saturação de disco.",
                    storage,
                    [
                        "Analisar maiores diretórios",
                        "Limpar temporários com segurança",
                        "Validar perfil do usuário",
                    ],
                )
            )
        elif storage:
            diagnoses.append(
                Diagnosis(
                    "STORAGE_DEGRADATION",
                    "Possível degradação por armazenamento",
                    "média",
                    "Há indicadores relevantes de espaço ou atividade de disco.",
                    storage,
                    ["Executar diagnóstico de armazenamento"],
                )
            )

        performance = [
            by[key]
            for key in ("CPU_PRESSURE", "MEMORY_PRESSURE")
            if key in by
        ]
        if len(performance) >= 2:
            diagnoses.append(
                Diagnosis(
                    "RESOURCE_PRESSURE",
                    "Pressão de CPU e memória",
                    "alta",
                    "CPU e RAM encontram-se simultaneamente elevadas.",
                    performance,
                    [
                        "Coletar amostragem detalhada",
                        "Analisar processos dominantes",
                    ],
                )
            )
        elif performance:
            diagnoses.append(
                Diagnosis(
                    "RESOURCE_PRESSURE_PARTIAL",
                    "Pressão de recursos",
                    "média",
                    performance[0].message,
                    performance,
                    ["Analisar processos dominantes"],
                )
            )

        security = [
            by[key]
            for key in ("DEFENDER_DISABLED", "FIREWALL_DISABLED")
            if key in by
        ]
        if security:
            diagnoses.append(
                Diagnosis(
                    "SECURITY_POSTURE",
                    "Desvio de postura de segurança",
                    "alta",
                    "Controles esperados estão desabilitados.",
                    security,
                    ["Validar política e GPO antes de remediar"],
                )
            )

        groups = (
            (
                "NETWORK_CONFIGURATION",
                "Falha provável de configuração de rede",
                ("NETWORK_NO_IPV4", "NETWORK_NO_GATEWAY", "NETWORK_NO_DNS", "NETWORK_NO_UP_ADAPTER"),
                "Há evidências objetivas de configuração/conectividade local incompleta.",
                ["Validar adaptador, DHCP, gateway e DNS"],
            ),
            (
                "PRINTING_FAILURE",
                "Falha provável no subsistema de impressão",
                ("SPOOLER_STOPPED", "PRINT_QUEUE_PENDING", "PRINTER_NONE"),
                "O playbook de impressão encontrou desvios relevantes.",
                ["Validar Spooler, fila, driver e conectividade da impressora"],
            ),
            (
                "DOMAIN_FAILURE",
                "Falha provável de domínio/GPO",
                ("DOMAIN_SECURE_CHANNEL_BROKEN", "DOMAIN_CONTROLLER_UNRESOLVED"),
                "A estação apresenta desvio de secure channel ou acesso ao DC.",
                ["Validar DNS, horário, DC e secure channel"],
            ),
            (
                "WINDOWS_UPDATE_STATE",
                "Windows Update requer atenção",
                ("UPDATE_QUERY_FAILED", "UPDATE_REBOOT_REQUIRED", "UPDATE_PENDING"),
                "O playbook encontrou pendência ou falha no Windows Update.",
                ["Validar serviços, componentes e reinicialização"],
            ),
            (
                "APPLICATION_STABILITY",
                "Instabilidade recente de aplicação",
                ("APP_CRASHES_RECENT",),
                "Há eventos recentes de crash de aplicações.",
                ["Correlacionar módulo, versão e horário do crash"],
            ),
            (
                "BSOD_HISTORY",
                "Evidência de BSOD",
                ("BSOD_EVIDENCE",),
                "Há bugchecks ou dumps disponíveis para análise.",
                ["Preservar e analisar os dumps"],
            ),
            (
                "GLPI_AGENT_FAILURE",
                "Falha provável do GLPI Agent",
                ("GLPI_NOT_INSTALLED", "GLPI_STOPPED"),
                "O agente não está instalado ou não está em execução.",
                ["Consultar log, validar serviço e política de instalação"],
            ),
        )

        for code, title, keys, rationale, actions in groups:
            evidence = [by[key] for key in keys if key in by]
            if evidence:
                diagnoses.append(
                    Diagnosis(
                        code,
                        title,
                        "alta" if any(f.severity.value in {"high", "critical"} for f in evidence) else "média",
                        rationale,
                        evidence,
                        actions,
                    )
                )

        return diagnoses
