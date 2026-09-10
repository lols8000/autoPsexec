from __future__ import annotations

import base64
import ctypes
import json
import locale
import os
import queue
import shutil
import socket
import subprocess
import threading
import time
from pathlib import Path
from typing import Callable, Iterable, Literal

from .host_identity import HostIdentity
from .logger import AuditLogger
from .result import CommandResult
from .retry import RetryPolicy
from .transport import LocalTransport, PsExecTransport, TransportManager, WinRMTransport

FallbackMode = Literal["read_only", "mutation", "none"]


class RemoteExecutor:
    """Fachada única de execução Local → WinRM → PsExec.

    Leituras podem fazer fallback após falha de transporte. Mutações só fazem
    fallback quando a falha é comprovadamente anterior à execução remota.
    """

    def __init__(
        self,
        *,
        psexec_path: str | None = None,
        timeout: int = 60,
        logger: AuditLogger | None = None,
        transport_cache_ttl_seconds: float = 120.0,
        retry_attempts: int = 2,
        retry_base_delay_seconds: float = 0.5,
    ) -> None:
        self.timeout = timeout
        self.logger = logger

        configured = Path(psexec_path) if psexec_path else None
        self.psexec_path = (
            str(configured)
            if configured and configured.exists()
            else self._discover_psexec()
        )

        self.retry_policy = RetryPolicy(
            max_attempts=int(retry_attempts),
            base_delay_seconds=float(retry_base_delay_seconds),
        )
        self.local_transport = LocalTransport(
            self._run_local,
            self._powershell_utf8_prefix,
        )
        self.winrm_transport = WinRMTransport(
            self._run_local,
            self._powershell_utf8_prefix,
        )
        self.psexec_transport = PsExecTransport(
            self._run_local,
            self._powershell_utf8_prefix,
            self.psexec_path,
        )
        self.transport_manager = TransportManager(
            self.local_transport,
            self.winrm_transport,
            self.psexec_transport,
            cache_ttl_seconds=transport_cache_ttl_seconds,
            retry_policy=self.retry_policy,
        )

    @staticmethod
    def _discover_psexec() -> str | None:
        candidates = (
            shutil.which("PsExec.exe"),
            shutil.which("psexec.exe"),
            r"C:\Windows\System32\PsExec.exe",
            r"C:\Sysinternals\PsExec.exe",
        )
        for candidate in candidates:
            if candidate and Path(candidate).exists():
                return str(candidate)
        return None

    @staticmethod
    def _console_encoding() -> str | None:
        if os.name != "nt":
            return None
        try:
            code_page = int(ctypes.windll.kernel32.GetConsoleOutputCP())
            return f"cp{code_page}" if code_page > 0 else None
        except (AttributeError, OSError, ValueError):
            return None

    @classmethod
    def _decode_output(
        cls,
        data: bytes,
        *,
        preferred: str | None = None,
    ) -> str:
        if not data:
            return ""

        if data.startswith((b"\xff\xfe", b"\xfe\xff")):
            try:
                return data.decode("utf-16")
            except UnicodeDecodeError:
                pass

        if data.startswith(b"\xef\xbb\xbf"):
            try:
                return data.decode("utf-8-sig")
            except UnicodeDecodeError:
                pass

        candidates: list[str] = []
        for encoding in (
            preferred,
            "utf-8",
            cls._console_encoding(),
            locale.getpreferredencoding(False),
            "cp850",
            "cp1252",
        ):
            if encoding and encoding.casefold() not in {
                item.casefold() for item in candidates
            }:
                candidates.append(encoding)

        for encoding in candidates:
            try:
                return data.decode(encoding, errors="strict")
            except (LookupError, UnicodeDecodeError):
                continue

        return data.decode("latin-1", errors="strict")

    @staticmethod
    def _powershell_utf8_prefix() -> str:
        return (
            "$__centralN2Utf8=New-Object System.Text.UTF8Encoding($false);"
            "[Console]::OutputEncoding=$__centralN2Utf8;"
            "$OutputEncoding=$__centralN2Utf8;"
        )

    @staticmethod
    def resolve_host(host: str) -> str | None:
        try:
            return socket.gethostbyname(host)
        except OSError:
            return None

    @staticmethod
    def is_local(host: str) -> bool:
        return HostIdentity.is_local(host)

    def select_transport(
        self,
        host: str,
        *,
        refresh: bool = False,
        winrm_result: CommandResult | None = None,
        psexec_result: CommandResult | None = None,
    ) -> str:
        return self.transport_manager.select(
            host,
            refresh=refresh,
            winrm_result=winrm_result,
            psexec_result=psexec_result,
        ).name

    def invalidate_transport(self, host: str) -> None:
        self.transport_manager.invalidate(host)

    def ping(self, host: str) -> CommandResult:
        if self.is_local(host):
            return CommandResult(
                True,
                "local",
                host,
                stdout="OK",
                transport="local",
            )
        return self._run_local(
            ["ping", "-n", "1", "-w", "1200", host],
            host=host,
            action="ping",
        )

    def test_admin_share(self, host: str) -> CommandResult:
        if self.is_local(host):
            return CommandResult(
                True,
                "local",
                host,
                stdout="OK",
                transport="local",
            )
        return self._run_local(
            ["cmd.exe", "/d", "/c", f"dir \\\\{host}\\admin$ >nul 2>&1"],
            host=host,
            action="admin_share",
        )

    def test_winrm(self, host: str) -> CommandResult:
        return self.retry_policy.run(lambda: self.winrm_transport.test(host))

    def test_psexec(self, host: str) -> CommandResult:
        return self.psexec_transport.test(host)

    @staticmethod
    def _transport_failure_kind(result: CommandResult) -> str | None:
        if result.success:
            return None

        text = f"{result.stderr}\n{result.stdout}".casefold()
        pre_execution = (
            "cannotuseipaddress",
            "trustedhosts",
            "access is denied",
            "acesso negado",
            "logon failure",
            "falha de logon",
            "connection refused",
            "actively refused",
            "nenhuma conexão pôde ser feita",
            "cannot connect to the destination",
            "não pode se conectar ao destino",
        )
        if any(marker in text for marker in pre_execution):
            return "pre_execution"

        indeterminate = (
            "psremotingtransportexception",
            "pssessionstatebroken",
            "winrmoperationtimeout",
            "operation timeout",
            "timed out",
            "timeout",
            "connection reset",
            "ws-management",
            "wsman",
            "the client cannot connect",
            "o cliente não conseguiu se conectar",
        )
        if result.return_code == 124 or any(marker in text for marker in indeterminate):
            return "indeterminate"

        return None

    def _fallback_after_winrm_failure(
        self,
        *,
        host: str,
        result: CommandResult,
        mode: FallbackMode,
        fallback: Callable[[], CommandResult],
    ) -> CommandResult:
        kind = self._transport_failure_kind(result)
        if not kind:
            return result

        can_fallback = (
            self.psexec_transport.available()
            and mode != "none"
            and (mode == "read_only" or kind == "pre_execution")
        )
        if can_fallback:
            self.transport_manager.invalidate(host)
            replacement = fallback()
            replacement.metadata["fallback_from"] = "winrm"
            replacement.metadata["fallback_reason"] = kind
            return replacement

        if kind == "indeterminate":
            result.mark_indeterminate(
                "A comunicação WinRM foi perdida durante uma ação que pode ter "
                "chegado ao destino. Valide o estado antes de repetir."
            )
        result.metadata["fallback_suppressed"] = True
        result.metadata["transport_failure_kind"] = kind
        return result

    def execute_powershell(
        self,
        host: str,
        script: str,
        *,
        timeout: int | None = None,
        fallback_mode: FallbackMode = "read_only",
    ) -> CommandResult:
        transport = self.transport_manager.select(host)
        result = transport.execute_powershell(host, script, timeout=timeout)
        if transport.name != "winrm":
            return result

        return self._fallback_after_winrm_failure(
            host=host,
            result=result,
            mode=fallback_mode,
            fallback=lambda: self.psexec_transport.execute_powershell(
                host,
                script,
                timeout=timeout,
            ),
        )

    def execute_mutating_powershell(
        self,
        host: str,
        script: str,
        *,
        timeout: int | None = None,
    ) -> CommandResult:
        return self.execute_powershell(
            host,
            script,
            timeout=timeout,
            fallback_mode="mutation",
        )

    @staticmethod
    def _parse_json_output(text: str):
        """Extrai o primeiro JSON válido mesmo quando o transporte adiciona ruído."""
        if not text or not text.strip():
            return None

        value = text.strip()
        try:
            return json.loads(value)
        except json.JSONDecodeError:
            pass

        decoder = json.JSONDecoder()
        for index, char in enumerate(value):
            if char not in "[{":
                continue
            try:
                data, _ = decoder.raw_decode(value[index:])
                return data
            except json.JSONDecodeError:
                continue
        return None

    def execute_powershell_json(
        self,
        host: str,
        script: str,
        *,
        timeout: int | None = None,
    ) -> CommandResult:
        result = self.execute_powershell(
            host,
            f"$r=& {{ {script} }};$r|ConvertTo-Json -Depth 8 -Compress",
            timeout=timeout,
            fallback_mode="read_only",
        )
        if result.success and result.stdout.strip():
            parsed = self._parse_json_output(result.stdout)
            if parsed is not None:
                result.data = parsed
            else:
                result.metadata["json_parse_error"] = True
        return result

    def execute_mutating_powershell_json(
        self,
        host: str,
        script: str,
        *,
        timeout: int | None = None,
    ) -> CommandResult:
        result = self.execute_powershell(
            host,
            f"$r=& {{ {script} }};$r|ConvertTo-Json -Depth 8 -Compress",
            timeout=timeout,
            fallback_mode="mutation",
        )
        if result.success and result.stdout.strip():
            parsed = self._parse_json_output(result.stdout)
            if parsed is not None:
                result.data = parsed
            else:
                result.metadata["json_parse_error"] = True
                result.mark_indeterminate(
                    "A ação terminou, mas o resultado estruturado não pôde ser validado."
                )
        return result

    def execute_psexec(
        self,
        host: str,
        executable: str,
        args: Iterable[str] = (),
        *,
        system: bool = False,
        timeout: int | None = None,
        output_encoding: str | None = None,
    ) -> CommandResult:
        return self.psexec_transport.execute_raw(
            host,
            executable,
            list(args),
            system=system,
            timeout=timeout,
            output_encoding=output_encoding,
        )

    def execute_cmd(
        self,
        host: str,
        command: str,
        *,
        timeout: int | None = None,
        fallback_mode: FallbackMode = "read_only",
    ) -> CommandResult:
        transport = self.transport_manager.select(host)
        result = transport.execute_cmd(host, command, timeout=timeout)
        if transport.name != "winrm":
            return result

        return self._fallback_after_winrm_failure(
            host=host,
            result=result,
            mode=fallback_mode,
            fallback=lambda: self.psexec_transport.execute_cmd(
                host,
                command,
                timeout=timeout,
            ),
        )

    def execute_mutating_cmd(
        self,
        host: str,
        command: str,
        *,
        timeout: int | None = None,
    ) -> CommandResult:
        return self.execute_cmd(
            host,
            command,
            timeout=timeout,
            fallback_mode="mutation",
        )

    def execute_remote_powershell_with_fallback(
        self,
        host: str,
        script: str,
        *,
        timeout: int | None = None,
    ) -> CommandResult:
        """Compatibilidade: use apenas para leitura.

        Mutações devem chamar execute_mutating_powershell().
        """
        return self.execute_powershell(
            host,
            script,
            timeout=timeout,
            fallback_mode="read_only",
        )

    def execute_cmd_stream(
        self,
        host: str,
        command: str,
        *,
        timeout: int | None = None,
        on_line: Callable[[str], None] | None = None,
    ) -> CommandResult:
        transport = self.transport_manager.select(host)
        remote_exit_marker = "__CENTRAL_N2_EXIT_CODE__="

        if transport.name == "local":
            cmd = ["cmd.exe", "/d", "/c", command]
            encoding = None
        elif transport.name == "winrm":
            safe_host = host.replace("'", "''")
            encoded_command = base64.b64encode(
                command.encode("utf-16le")
            ).decode("ascii")
            payload = self._powershell_utf8_prefix() + f"""
$ErrorActionPreference='Stop'
$__command=[Text.Encoding]::Unicode.GetString(
    [Convert]::FromBase64String('{encoded_command}')
)
Invoke-Command -ComputerName '{safe_host}' -ScriptBlock {{
    param([string]$Command)
    & cmd.exe /d /c $Command
    Write-Output '{remote_exit_marker}'$LASTEXITCODE
}} -ArgumentList $__command
"""
            cmd = [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                payload,
            ]
            encoding = "utf-8"
        else:
            cmd = [
                str(self.psexec_path),
                "-accepteula",
                "-nobanner",
                f"\\\\{host}",
                "cmd.exe",
                "/d",
                "/c",
                command,
            ]
            encoding = None

        def filtered_line(line: str) -> None:
            if line.startswith(remote_exit_marker):
                return
            if on_line:
                on_line(line)

        result = self._run_local_streaming(
            cmd,
            host=host,
            action="cmd_stream",
            timeout=timeout,
            output_encoding=encoding,
            on_line=filtered_line,
        )
        result.transport = transport.name

        if transport.name == "winrm" and result.success:
            lines = result.stdout.splitlines()
            marker_lines = [
                line for line in lines if line.strip().startswith(remote_exit_marker)
            ]
            result.stdout = "\n".join(
                line for line in lines
                if not line.strip().startswith(remote_exit_marker)
            ).strip()
            if marker_lines:
                try:
                    remote_code = int(
                        marker_lines[-1].strip().split("=", 1)[1]
                    )
                    result.return_code = remote_code
                    result.success = remote_code == 0
                    result.metadata["remote_exit_code_confirmed"] = True
                    if remote_code != 0:
                        result.stderr = (
                            f"Comando remoto retornou exit code {remote_code}."
                        )
                except (ValueError, IndexError):
                    result.mark_indeterminate(
                        "A execução terminou, mas o exit code remoto não pôde ser validado."
                    )
            else:
                result.mark_indeterminate(
                    "A execução terminou sem o marcador de exit code remoto."
                )

        if result.return_code == 124 and transport.name != "local":
            result.mark_indeterminate(
                "Timeout local: o processo remoto pode continuar em execução."
            )

        return result

    def _run_powershell_local(
        self,
        script: str,
        *,
        host: str,
        action: str,
        timeout: int | None = None,
    ) -> CommandResult:
        return self._run_local(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                self._powershell_utf8_prefix() + script,
            ],
            host=host,
            action=action,
            timeout=timeout,
            output_encoding="utf-8",
        )

    def _run_local_streaming(
        self,
        cmd: list[str],
        *,
        host: str,
        action: str,
        timeout: int | None = None,
        output_encoding: str | None = None,
        on_line: Callable[[str], None] | None = None,
    ) -> CommandResult:
        started = time.perf_counter()
        printable = subprocess.list2cmdline(cmd)
        lines: list[str] = []
        process: subprocess.Popen[bytes] | None = None
        reader_thread: threading.Thread | None = None

        try:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                shell=False,
            )
            assert process.stdout is not None

            events: queue.Queue[bytes | None] = queue.Queue()

            def reader() -> None:
                try:
                    for raw_line in iter(process.stdout.readline, b""):
                        events.put(raw_line)
                finally:
                    events.put(None)

            reader_thread = threading.Thread(
                target=reader,
                name="central-n2-stream-reader",
                daemon=True,
            )
            reader_thread.start()

            effective_timeout = timeout or self.timeout
            deadline = time.monotonic() + effective_timeout
            reader_done = False

            while not reader_done:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    process.kill()
                    process.wait(timeout=5)
                    raise subprocess.TimeoutExpired(cmd, effective_timeout)

                try:
                    raw = events.get(timeout=min(0.1, remaining))
                except queue.Empty:
                    continue

                if raw is None:
                    reader_done = True
                    continue

                text = self._decode_output(
                    raw,
                    preferred=output_encoding,
                ).rstrip()
                lines.append(text)
                if on_line:
                    on_line(text)

            if reader_thread:
                reader_thread.join(timeout=1)

            remaining = max(0.1, deadline - time.monotonic())
            return_code = process.wait(timeout=remaining)
            result = CommandResult(
                return_code == 0,
                printable,
                host,
                stdout="\n".join(lines),
                return_code=return_code,
                duration_ms=int((time.perf_counter() - started) * 1000),
            )
        except subprocess.TimeoutExpired:
            if process and process.poll() is None:
                process.kill()
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    pass
            result = CommandResult.failure(
                host,
                printable,
                f"Timeout após {timeout or self.timeout}s",
                return_code=124,
            )
            result.duration_ms = int((time.perf_counter() - started) * 1000)
        except OSError as exc:
            result = CommandResult.failure(
                host,
                printable,
                str(exc),
                return_code=127,
            )
            result.duration_ms = int((time.perf_counter() - started) * 1000)

        if self.logger:
            self.logger.log_result(action, result)
        return result

    def _run_local(
        self,
        cmd: list[str],
        *,
        host: str,
        action: str,
        timeout: int | None = None,
        output_encoding: str | None = None,
    ) -> CommandResult:
        started = time.perf_counter()
        printable = subprocess.list2cmdline(cmd)

        try:
            process = subprocess.run(
                cmd,
                capture_output=True,
                text=False,
                timeout=timeout or self.timeout,
                shell=False,
            )
            result = CommandResult(
                process.returncode == 0,
                printable,
                host,
                stdout=self._decode_output(
                    process.stdout,
                    preferred=output_encoding,
                ).strip(),
                stderr=self._decode_output(
                    process.stderr,
                    preferred=output_encoding,
                ).strip(),
                return_code=process.returncode,
                duration_ms=int((time.perf_counter() - started) * 1000),
            )
        except subprocess.TimeoutExpired:
            result = CommandResult.failure(
                host,
                printable,
                f"Timeout após {timeout or self.timeout}s",
                return_code=124,
            )
            result.duration_ms = int((time.perf_counter() - started) * 1000)
        except OSError as exc:
            result = CommandResult.failure(
                host,
                printable,
                str(exc),
                return_code=127,
            )
            result.duration_ms = int((time.perf_counter() - started) * 1000)

        if self.logger:
            self.logger.log_result(action, result)
        return result
