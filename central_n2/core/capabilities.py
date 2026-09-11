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

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
$isAdmin = $principal.IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)

$winget = Get-Command winget.exe -ErrorAction SilentlyContinue
$defender = Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue
$bitlocker = Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue
$tpm = Get-Command Get-Tpm -ErrorAction SilentlyContinue
$physical = Get-Command Get-PhysicalDisk -ErrorAction SilentlyContinue
$netAdapter = Get-Command Get-NetAdapter -ErrorAction SilentlyContinue
$pnp = Get-Command Get-PnpDevice -ErrorAction SilentlyContinue
$printer = Get-Command Get-Printer -ErrorAction SilentlyContinue
$certImport = Get-Command Import-Certificate -ErrorAction SilentlyContinue

$secure = $null
try { $secure = Confirm-SecureBootUEFI -ErrorAction Stop } catch {}

$windowsUpdateCom = $false
try {
    $wu = New-Object -ComObject Microsoft.Update.Session
    $windowsUpdateCom = [bool]$wu
} catch {}

[pscustomobject]@{
  PowerShellVersion=$PSVersionTable.PSVersion.ToString()
  OS=$os.Caption
  Build=$os.BuildNumber
  Architecture=$os.OSArchitecture
  Manufacturer=$cs.Manufacturer
  Model=$cs.Model
  DomainMember=[bool]$cs.PartOfDomain
  ExecutionIdentity=$identity.Name
  IsAdmin=[bool]$isAdmin
  IsSystem=[bool]$identity.IsSystem
  Winget=[bool]$winget
  Defender=[bool]$defender
  BitLocker=[bool]$bitlocker
  TPM=[bool]$tpm
  PhysicalDisk=[bool]$physical
  NetAdapter=[bool]$netAdapter
  PnpDevice=[bool]$pnp
  PrinterManagement=[bool]$printer
  CertificateImport=[bool]$certImport
  PnPUtil=[bool](Test-Path "$env:SystemRoot\System32\pnputil.exe")
  DISM=[bool](Test-Path "$env:SystemRoot\System32\dism.exe")
  SFC=[bool](Test-Path "$env:SystemRoot\System32\sfc.exe")
  UsoClient=[bool](Test-Path "$env:SystemRoot\System32\UsoClient.exe")
  Klist=[bool](Test-Path "$env:SystemRoot\System32\klist.exe")
  W32Time=[bool](Get-Service w32time -ErrorAction SilentlyContinue)
  WindowsUpdateCOM=[bool]$windowsUpdateCom
  SecureBoot=$secure
  Battery=($bat.Count -gt 0)
  GLPI=[bool](
      Get-Service -ErrorAction SilentlyContinue |
      Where-Object {$_.Name -match 'glpi'} |
      Select-Object -First 1
  )
}
'''
        result = self.executor.execute_powershell_json(
            host,
            script,
            timeout=90,
        )
        if not result.success:
            return CapabilityReport(
                host,
                result.transport,
                error=result.stderr,
            )
        values = (
            result.data
            if isinstance(result.data, dict)
            else {}
        )
        return CapabilityReport(
            host,
            result.transport,
            values=values,
        )
