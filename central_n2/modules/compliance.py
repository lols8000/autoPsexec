from __future__ import annotations
from dataclasses import dataclass,asdict
from typing import Any
@dataclass(frozen=True)
class ComplianceItem:key:str;label:str;compliant:bool;actual:Any;expected:Any;severity:str="medium"
def evaluate_compliance(snapshot:dict,baseline:dict)->dict:
    items=[]
    def add(key,label,ok,actual,expected,severity="medium"):items.append(ComplianceItem(key,label,ok,actual,expected,severity))
    minimum=float(baseline.get("min_disk_free_percent",15));add("disk","Espaço livre em C:",float(snapshot.get("DiskFreePercent") or 0)>=minimum,snapshot.get("DiskFreePercent"),f">= {minimum}%","high")
    if baseline.get("defender_required",True):add("defender","Microsoft Defender",snapshot.get("DefenderEnabled") is True,snapshot.get("DefenderEnabled"),True,"critical")
    if baseline.get("firewall_required",True):add("firewall","Firewall",snapshot.get("FirewallEnabled") is True,snapshot.get("FirewallEnabled"),True,"high")
    if baseline.get("glpi_required",True):add("glpi","GLPI Agent",snapshot.get("GlpiRunning") is True,snapshot.get("GlpiRunning"),True,"medium")
    maximum=int(baseline.get("max_uptime_days",30));add("uptime","Uptime",int(snapshot.get("UptimeDays") or 0)<=maximum,snapshot.get("UptimeDays"),f"<= {maximum} dias","low")
    if baseline.get("pending_reboot_not_allowed",True):add("pending_reboot","Reinicialização pendente",snapshot.get("PendingReboot") is not True,snapshot.get("PendingReboot"),False,"medium")
    if baseline.get("bitlocker_required",False):add("bitlocker","BitLocker",snapshot.get("BitLockerProtected") is True,snapshot.get("BitLockerProtected"),True,"high")
    if baseline.get("tpm_required",False):add("tpm","TPM pronto",snapshot.get("TPMReady") is True,snapshot.get("TPMReady"),True,"high")
    if baseline.get("secure_boot_required",False):add("secure_boot","Secure Boot",snapshot.get("SecureBoot") is True,snapshot.get("SecureBoot"),True,"high")
    compliant=sum(1 for x in items if x.compliant);score=round((compliant/len(items))*100) if items else 100
    return {"score":score,"compliant":compliant,"total":len(items),"items":[asdict(x) for x in items]}
