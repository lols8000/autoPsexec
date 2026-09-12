from __future__ import annotations

from .models import QualificationCase


BASELINE_CASES = (
    QualificationCase(
        "core.connectivity",
        "Conectividade e transporte administrativo",
        "Core",
        "session",
    ),
    QualificationCase(
        "core.capabilities",
        "Capabilities do endpoint",
        "Core",
        "capabilities",
    ),
    QualificationCase(
        "core.health",
        "Health snapshot",
        "Core",
        "probe",
    ),
    QualificationCase(
        "inventory.processes",
        "Inventário de processos",
        "Inventário",
        "probe",
    ),
    QualificationCase(
        "inventory.services",
        "Inventário de serviços",
        "Inventário",
        "probe",
    ),
    QualificationCase(
        "network.adapters",
        "Adaptadores de rede",
        "Rede",
        "probe",
    ),
    QualificationCase(
        "network.ip",
        "Configuração IP / Gateway / DNS",
        "Rede",
        "probe",
    ),
    QualificationCase(
        "devices.problems",
        "Dispositivos PnP com problema",
        "Hardware",
        "probe",
    ),
    QualificationCase(
        "security.posture",
        "Postura de segurança",
        "Segurança",
        "probe",
    ),
    QualificationCase(
        "domain.status",
        "Domínio / secure channel / horário",
        "Domínio",
        "probe",
        required=False,
        domain_member=True,
    ),
    QualificationCase(
        "printers.inventory",
        "Inventário de impressoras",
        "Impressão",
        "probe",
        required=False,
        capability="PrinterManagement",
    ),
    QualificationCase(
        "glpi.status",
        "GLPI Agent",
        "GLPI",
        "probe",
        required=False,
        capability="GLPI",
    ),
    QualificationCase(
        "updates.status",
        "Windows Update",
        "Windows Update",
        "probe",
        required=False,
        capability="WindowsUpdateCOM",
    ),
)


RECOMMENDED_ACTIONS = {
    "safe": (
        "network.flush_dns",
        "printer.restart_spooler",
        "glpi.force_inventory",
        "domain.gpupdate",
        "defender.signatures",
    ),
    "recovery": (
        "network.renew_dhcp",
        "network.adapter_restart",
        "energy.restart",
    ),
}


PROFILES = {
    "auto": "Detecta capacidades e pula casos não aplicáveis.",
    "local": "Endpoint local administrativo.",
    "winrm": "Endpoint remoto com WinRM utilizável.",
    "psexec": "Endpoint remoto com ADMIN$/PsExec e WinRM indisponível.",
    "notebook": "Notebook/Wi-Fi, incluindo testes de recovery quando autorizados.",
    "printer": "Estação com PrintManagement/Spooler.",
    "glpi": "Estação com GLPI Agent.",
    "domain": "Estação membro de domínio.",
}
