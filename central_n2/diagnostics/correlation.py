from __future__ import annotations
from .models import Diagnosis, Finding

class CorrelationEngine:
    def correlate(self, findings: list[Finding]) -> list[Diagnosis]:
        by={f.id:f for f in findings}; out=[]
        storage=[by[k] for k in ("DISK_CRITICAL","DISK_LOW","DISK_SATURATION") if k in by]
        if "DISK_CRITICAL" in by and "DISK_SATURATION" in by:
            out.append(Diagnosis("STORAGE_PRESSURE","Pressão de armazenamento","alta","Espaço livre crítico combinado com saturação de disco.",storage,["Analisar maiores diretórios","Limpar temporários com segurança","Validar perfil do usuário"]))
        elif storage:
            out.append(Diagnosis("STORAGE_DEGRADATION","Possível degradação por armazenamento","média","Há indicadores relevantes de espaço ou atividade de disco.",storage,["Executar diagnóstico de armazenamento"]))
        perf=[by[k] for k in ("CPU_PRESSURE","MEMORY_PRESSURE") if k in by]
        if len(perf)>=2:
            out.append(Diagnosis("RESOURCE_PRESSURE","Pressão de CPU e memória","alta","CPU e RAM encontram-se simultaneamente elevadas.",perf,["Coletar amostragem detalhada","Analisar processos dominantes"]))
        sec=[by[k] for k in ("DEFENDER_DISABLED","FIREWALL_DISABLED") if k in by]
        if sec:
            out.append(Diagnosis("SECURITY_POSTURE","Desvio de postura de segurança","alta","Controles esperados estão desabilitados.",sec,["Validar política e GPO antes de remediar"]))
        if "GLPI_STOPPED" in by:
            out.append(Diagnosis("GLPI_AGENT_FAILURE","Falha provável do GLPI Agent","média","O agente não está em execução.",[by["GLPI_STOPPED"]],["Consultar log","Validar serviço","Forçar inventário"]))
        return out
