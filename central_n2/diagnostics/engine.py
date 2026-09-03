from __future__ import annotations
from typing import Any
from .models import Finding, Severity

class DiagnosticEngine:
    def evaluate(self, data: dict[str, Any]) -> list[Finding]:
        out=[]
        disk=self._num(data,"DiskFreePercent","disk_free_percent","FreePercent")
        ram=self._num(data,"RAMUsedPercent","ram_percent","MemoryPercent")
        cpu=self._num(data,"CPUPercent","cpu_percent","CpuAverage")
        active=self._num(data,"DiskActivePercent","disk_active_percent","DiskAverage")
        if disk is not None and disk < 5:
            out.append(Finding("DISK_CRITICAL",Severity.CRITICAL,f"Disco com apenas {disk:.1f}% livre",{"free_percent":disk},["Executar análise de espaço","Executar limpeza segura"]))
        elif disk is not None and disk < 15:
            out.append(Finding("DISK_LOW",Severity.HIGH,f"Espaço livre baixo: {disk:.1f}%",{"free_percent":disk},["Analisar perfis e temporários"]))
        if ram is not None and ram >= 90:
            out.append(Finding("MEMORY_PRESSURE",Severity.HIGH,f"Uso de RAM em {ram:.1f}%",{"ram_percent":ram},["Analisar top processos por memória"]))
        if cpu is not None and cpu >= 90:
            out.append(Finding("CPU_PRESSURE",Severity.HIGH,f"CPU em {cpu:.1f}%",{"cpu_percent":cpu},["Analisar top processos por CPU"]))
        if active is not None and active >= 90:
            out.append(Finding("DISK_SATURATION",Severity.HIGH,f"Disco ativo em {active:.1f}%",{"disk_active_percent":active},["Correlacionar com espaço livre e processos"]))
        if bool(data.get("PendingReboot")):
            out.append(Finding("PENDING_REBOOT",Severity.MEDIUM,"Reinicialização pendente",recommendations=["Validar janela e reiniciar"]))
        stopped=int(data.get("StoppedAutoServices") or 0)
        if stopped:
            out.append(Finding("AUTO_SERVICES_STOPPED",Severity.HIGH,f"{stopped} serviço(s) automático(s) parado(s)",{"count":stopped},["Inspecionar serviços e eventos"]))
        if data.get("DefenderEnabled") is False:
            out.append(Finding("DEFENDER_DISABLED",Severity.CRITICAL,"Microsoft Defender desabilitado"))
        if data.get("FirewallEnabled") is False:
            out.append(Finding("FIREWALL_DISABLED",Severity.HIGH,"Firewall do Windows desabilitado"))
        if data.get("GlpiRunning") is False:
            out.append(Finding("GLPI_STOPPED",Severity.MEDIUM,"GLPI Agent não está em execução",recommendations=["Validar serviço e log do agente"]))
        return out

    @staticmethod
    def _num(data: dict[str, Any], *keys: str) -> float | None:
        for key in keys:
            try:
                if data.get(key) is not None: return float(data[key])
            except (TypeError,ValueError): pass
        return None
