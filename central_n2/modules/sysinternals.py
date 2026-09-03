from __future__ import annotations

from pathlib import PureWindowsPath

from core.executor import RemoteExecutor
from core.result import CommandResult


class SysinternalsModule:
    """Integra ferramentas Sysinternals quando já instaladas na estação alvo."""

    def __init__(self, executor: RemoteExecutor, tools_dir: str = r"C:\Sysinternals") -> None:
        self.executor = executor
        self.tools_dir = tools_dir.rstrip("\\")

    def inventory(self, host: str) -> CommandResult:
        root = self.tools_dir.replace("'", "''")
        script = f'''
$root='{root}'
$names='autorunsc.exe','procdump.exe','handle.exe','sigcheck.exe','psping.exe','psloggedon.exe','rammap.exe','procmon.exe'
$names | ForEach-Object {{
 $p=Join-Path $root $_
 [pscustomobject]@{{Name=$_;Path=$p;Available=(Test-Path $p)}}
}}
'''
        return self.executor.execute_powershell_json(host, script)

    def autoruns(self, host: str) -> CommandResult:
        exe = str(PureWindowsPath(self.tools_dir) / "autorunsc.exe")
        return self.executor.execute_cmd(host, f'"{exe}" -accepteula -a * -c -h -s -m', timeout=300)

    def capture_dump(self, host: str, process: str, dump_dir: str = r"C:\CentralN2\Dumps") -> CommandResult:
        safe_process = process.replace('"', '')
        exe = str(PureWindowsPath(self.tools_dir) / "procdump.exe")
        script = f'''
New-Item -ItemType Directory -Path '{dump_dir}' -Force | Out-Null
& '{exe}' -accepteula -ma '{safe_process}' '{dump_dir}'
'''
        return self.executor.execute_remote_powershell_with_fallback(host, script, timeout=600)

    def handle_search(self, host: str, text: str) -> CommandResult:
        safe = text.replace('"', '')
        exe = str(PureWindowsPath(self.tools_dir) / "handle.exe")
        return self.executor.execute_cmd(host, f'"{exe}" -accepteula "{safe}"', timeout=180)

    def sigcheck(self, host: str, path: str) -> CommandResult:
        safe = path.replace('"', '')
        exe = str(PureWindowsPath(self.tools_dir) / "sigcheck.exe")
        return self.executor.execute_cmd(host, f'"{exe}" -accepteula -h -i -q "{safe}"', timeout=180)
