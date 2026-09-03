from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class CapabilityReport:
    host: str
    transport: str
    values: dict[str, Any] = field(default_factory=dict)
    error: str | None = None


class CapabilityDetector:
    def __init__(self, executor) -> None:
        self.executor = executor

    def probe(self, host: str) -> CapabilityReport:
        script = r'''
$os = Get-CimInstance Win32_OperatingSystem
$cs = Get-CimInstance Win32_ComputerSystem
$bat = @(Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue)
$winget = Get-Command winget.exe -ErrorAction SilentlyContinue
$defender = Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue
$bitlocker = Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue
$tpm = Get-Command Get-Tpm -ErrorAction SilentlyContinue
$physical = Get-Command Get-PhysicalDisk -ErrorAction SilentlyContinue
$secure = $null
try { $secure = Confirm-SecureBootUEFI -ErrorAction Stop } catch {}
[pscustomobject]@{
  PowerShellVersion=$PSVersionTable.PSVersion.ToString()
  OS=$os.Caption
  Build=$os.BuildNumber
  Architecture=$os.OSArchitecture
  Manufacturer=$cs.Manufacturer
  Model=$cs.Model
  Winget=[bool]$winget
  Defender=[bool]$defender
  BitLocker=[bool]$bitlocker
  TPM=[bool]$tpm
  PhysicalDisk=[bool]$physical
  SecureBoot=$secure
  Battery=($bat.Count -gt 0)
  GLPI=[bool](Get-Service -ErrorAction SilentlyContinue | Where-Object {$_.Name -match 'glpi'} | Select-Object -First 1)
}
'''
        result = self.executor.execute_powershell_json(host, script, timeout=90)
        if not result.success:
            return CapabilityReport(host, result.transport, error=result.stderr)
        values = result.data if isinstance(result.data, dict) else {}
        return CapabilityReport(host, result.transport, values=values)
