from __future__ import annotations

from pathlib import PureWindowsPath

from core.executor import RemoteExecutor
from core.result import CommandResult
from core.validation import (
    quote_cmd_argument,
    quote_powershell_literal,
    validate_process_name,
    validate_windows_path,
)


class SysinternalsModule:
    """Integra ferramentas Sysinternals homologadas na estação alvo."""

    def __init__(
        self,
        executor: RemoteExecutor,
        tools_dir: str = r"C:\Sysinternals",
    ) -> None:
        self.executor = executor
        self.tools_dir = tools_dir.rstrip("\\")

    def _tool(self, name: str) -> str:
        return str(PureWindowsPath(self.tools_dir) / name)

    def inventory(self, host: str) -> CommandResult:
        root = quote_powershell_literal(self.tools_dir)
        script = f"""
$root={root}
$names=@(
    'autorunsc.exe','procdump.exe','handle.exe','sigcheck.exe',
    'psping.exe','psloggedon.exe','rammap.exe','procmon.exe'
)
$names | ForEach-Object {{
    $path=Join-Path $root $_
    [pscustomobject]@{{
        Name=$_
        Path=$path
        Available=(Test-Path $path)
    }}
}}
"""
        return self.executor.execute_powershell_json(host, script)

    def autoruns(self, host: str) -> CommandResult:
        exe = quote_cmd_argument(self._tool("autorunsc.exe"))
        return self.executor.execute_cmd(
            host,
            f"{exe} -accepteula -a * -c -h -s -m",
            timeout=300,
        )

    def capture_dump(
        self,
        host: str,
        process: str,
        dump_dir: str = r"C:\CentralN2\Dumps",
    ) -> CommandResult:
        safe_process = quote_powershell_literal(
            validate_process_name(process)
        )
        safe_dump_dir = quote_powershell_literal(
            validate_windows_path(dump_dir)
        )
        exe = quote_powershell_literal(self._tool("procdump.exe"))
        script = f"""
New-Item -ItemType Directory -Path {safe_dump_dir} -Force | Out-Null
& {exe} -accepteula -ma {safe_process} {safe_dump_dir}
if ($LASTEXITCODE -ne 0) {{
    throw "ProcDump retornou exit code $LASTEXITCODE"
}}
"""
        return self.executor.execute_mutating_powershell(
            host,
            script,
            timeout=600,
        )

    def handle_search(self, host: str, text: str) -> CommandResult:
        safe = quote_cmd_argument(text)
        exe = quote_cmd_argument(self._tool("handle.exe"))
        return self.executor.execute_cmd(
            host,
            f"{exe} -accepteula {safe}",
            timeout=180,
        )

    def sigcheck(self, host: str, path: str) -> CommandResult:
        safe = quote_cmd_argument(validate_windows_path(path))
        exe = quote_cmd_argument(self._tool("sigcheck.exe"))
        return self.executor.execute_cmd(
            host,
            f"{exe} -accepteula -h -i -q {safe}",
            timeout=180,
        )
