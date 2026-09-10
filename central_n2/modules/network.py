from __future__ import annotations

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import quote_powershell_literal, validate_port


class NetworkModule:
    def __init__(self, executor: RemoteExecutor) -> None:
        self.executor = executor

    def adapters(self, host: str) -> CommandResult:
        script = """
Get-NetAdapter |
    Sort-Object ifIndex |
    Select-Object Name,InterfaceDescription,Status,MacAddress,LinkSpeed,ifIndex
"""
        return self.executor.execute_powershell_json(host, script)

    def ip_configuration(self, host: str) -> CommandResult:
        script = """
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
        )

    def flush_dns(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_powershell(
            host,
            "Clear-DnsClientCache -ErrorAction Stop",
        )

    def reset_winsock(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(host, "netsh winsock reset")

    def reset_tcpip(self, host: str) -> CommandResult:
        return self.executor.execute_mutating_cmd(host, "netsh int ip reset")

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
        return self.executor.execute_mutating_powershell(host, script)

    def arp_table(self, host: str) -> CommandResult:
        script = """
Get-NetNeighbor |
    Sort-Object InterfaceIndex,IPAddress |
    Select-Object InterfaceIndex,IPAddress,LinkLayerAddress,State
"""
        return self.executor.execute_powershell_json(host, script)

    def connections(self, host: str) -> CommandResult:
        script = """
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
