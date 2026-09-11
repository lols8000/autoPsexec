from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import (
    quote_powershell_literal,
    validate_port,
    validate_safe_name,
)


class NetworkModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def adapters(self, host: str) -> CommandResult:
        script = r"""
Get-NetAdapter |
    Sort-Object ifIndex |
    Select-Object Name,InterfaceDescription,Status,MacAddress,LinkSpeed,ifIndex
"""
        return self.executor.execute_powershell_json(host, script)

    def adapter_status(
        self,
        host: str,
        adapter_name: str,
    ) -> CommandResult:
        safe = quote_powershell_literal(
            validate_safe_name(adapter_name, label="Adaptador")
        )
        script = f"""
$a = Get-NetAdapter -Name {safe} -ErrorAction Stop
[pscustomobject]@{{
    Name = $a.Name
    Status = $a.Status.ToString()
    InterfaceDescription = $a.InterfaceDescription
    ifIndex = $a.ifIndex
}}
"""
        return self.executor.execute_powershell_json(host, script)

    def ip_configuration(self, host: str) -> CommandResult:
        script = r"""
Get-NetIPConfiguration | ForEach-Object {
    [pscustomobject]@{
        InterfaceAlias = $_.InterfaceAlias
        IPv4 = ($_.IPv4Address.IPAddress -join ', ')
        Gateway = ($_.IPv4DefaultGateway.NextHop -join ', ')
        DNS = ($_.DNSServer.ServerAddresses -join ', ')
    }
}
"""
        return self.executor.execute_powershell_json(host, script)

    def renew_dhcp(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(
            host,
            "ipconfig /release && ipconfig /renew",
            timeout=180,
        )

    def flush_dns(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_powershell(
            host,
            "Clear-DnsClientCache -ErrorAction Stop",
        )

    def register_dns(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(
            host,
            "ipconfig /registerdns",
            timeout=120,
        )

    def reset_winsock(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(
            host,
            "netsh winsock reset",
        )

    def reset_tcpip(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(
            host,
            "netsh int ip reset",
        )

    def clear_arp(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(
            host,
            "arp -d *",
            timeout=60,
        )

    def set_adapter_state(
        self,
        host: str,
        adapter_name: str,
        *,
        enabled: bool,
    ) -> CommandResult:
        safe = quote_powershell_literal(
            validate_safe_name(adapter_name, label="Adaptador")
        )
        cmd = "Enable-NetAdapter" if enabled else "Disable-NetAdapter"
        script = f"""
{cmd} -Name {safe} -Confirm:$false -ErrorAction Stop
Start-Sleep -Milliseconds 500
$a = Get-NetAdapter -Name {safe} -ErrorAction Stop
[pscustomobject]@{{
    Name = $a.Name
    Status = $a.Status.ToString()
    Enabled = [bool]($a.Status -ne 'Disabled')
}}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=90,
        )

    def restart_adapter(
        self,
        host: str,
        adapter_name: str,
    ) -> CommandResult:
        safe = quote_powershell_literal(
            validate_safe_name(adapter_name, label="Adaptador")
        )
        script = f"""
$name = {safe}
$payload = @"
Start-Sleep -Seconds 2
Disable-NetAdapter -Name '$name' -Confirm:$false -ErrorAction Stop
Start-Sleep -Seconds 3
Enable-NetAdapter -Name '$name' -Confirm:$false -ErrorAction Stop
"@
$encoded = [Convert]::ToBase64String(
    [Text.Encoding]::Unicode.GetBytes($payload)
)
$p = Start-Process powershell.exe -ArgumentList @(
    '-NoProfile','-NonInteractive','-EncodedCommand',$encoded
) -WindowStyle Hidden -PassThru
[pscustomobject]@{{
    Scheduled = $true
    ProcessId = $p.Id
    Adapter = $name
}}
"""
        return self.executor.execute_mutating_powershell_json(
            host,
            script,
            timeout=60,
        )

    def wifi(self, host: str, enable: bool) -> CommandResult:
        action = "Enable-NetAdapter" if enable else "Disable-NetAdapter"
        script = f"""
$wifi = Get-NetAdapter | Where-Object {{
    $_.PhysicalMediaType -eq 'Native 802.11' -or
    $_.InterfaceDescription -match 'Wireless|Wi-Fi|802.11'
}}
if (-not $wifi) {{ throw 'Nenhum adaptador Wi-Fi encontrado.' }}
$wifi | {action} -Confirm:$false -ErrorAction Stop
$wifi | Select-Object Name,Status,MacAddress
"""
        return self.executor.execute_mutating_powershell_json(host, script)

    def arp_table(self, host: str) -> CommandResult:
        script = r"""
Get-NetNeighbor |
    Sort-Object InterfaceIndex,IPAddress |
    Select-Object InterfaceIndex,IPAddress,LinkLayerAddress,State
"""
        return self.executor.execute_powershell_json(host, script)

    def connections(self, host: str) -> CommandResult:
        script = r"""
Get-NetTCPConnection -State Established -ErrorAction SilentlyContinue |
    Select-Object LocalAddress,LocalPort,RemoteAddress,RemotePort,OwningProcess |
    Sort-Object OwningProcess
"""
        return self.executor.execute_powershell_json(host, script)

    def test_tcp(
        self,
        host: str,
        destination: str,
        port: int,
    ) -> CommandResult:
        safe_destination = quote_powershell_literal(destination)
        safe_port = validate_port(port)
        script = (
            f"Test-NetConnection -ComputerName {safe_destination} -Port {safe_port} | "
            "Select-Object ComputerName,RemoteAddress,RemotePort,TcpTestSucceeded"
        )
        return self.executor.execute_powershell_json(host, script)
