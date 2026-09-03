from __future__ import annotations
import ctypes,json,locale,os,queue,shutil,socket,subprocess,threading,time
from pathlib import Path
from typing import Iterable,Callable
from .host_identity import HostIdentity
from .logger import AuditLogger
from .result import CommandResult
from .transport import LocalTransport,PsExecTransport,TransportManager,WinRMTransport
class RemoteExecutor:
    def __init__(self,*,psexec_path=None,timeout=60,logger=None,transport_cache_ttl_seconds=120.0):
        self.timeout=timeout;self.logger=logger;configured=Path(psexec_path) if psexec_path else None;self.psexec_path=str(configured) if configured and configured.exists() else self._discover_psexec()
        self.local_transport=LocalTransport(self._run_local,self._powershell_utf8_prefix);self.winrm_transport=WinRMTransport(self._run_local,self._powershell_utf8_prefix);self.psexec_transport=PsExecTransport(self._run_local,self._powershell_utf8_prefix,self.psexec_path)
        self.transport_manager=TransportManager(self.local_transport,self.winrm_transport,self.psexec_transport,cache_ttl_seconds=transport_cache_ttl_seconds)
    @staticmethod
    def _discover_psexec():
        for c in (shutil.which("PsExec.exe"),shutil.which("psexec.exe"),r"C:\Windows\System32\PsExec.exe",r"C:\Sysinternals\PsExec.exe"):
            if c and Path(c).exists():return str(c)
        return None
    @staticmethod
    def _console_encoding():
        if os.name!="nt":return None
        try:cp=int(ctypes.windll.kernel32.GetConsoleOutputCP());return f"cp{cp}" if cp>0 else None
        except (AttributeError,OSError,ValueError):return None
    @classmethod
    def _decode_output(cls,data,*,preferred=None):
        if not data:return ""
        if data.startswith((b"\xff\xfe",b"\xfe\xff")):
            try:return data.decode("utf-16")
            except UnicodeDecodeError:pass
        if data.startswith(b"\xef\xbb\xbf"):
            try:return data.decode("utf-8-sig")
            except UnicodeDecodeError:pass
        candidates=[]
        for enc in (preferred,"utf-8",cls._console_encoding(),locale.getpreferredencoding(False),"cp850","cp1252"):
            if enc and enc.lower() not in {x.lower() for x in candidates}:candidates.append(enc)
        for enc in candidates:
            try:return data.decode(enc,errors="strict")
            except (LookupError,UnicodeDecodeError):pass
        return data.decode("latin-1",errors="strict")
    @staticmethod
    def _powershell_utf8_prefix():return "$__centralN2Utf8=New-Object System.Text.UTF8Encoding($false);[Console]::OutputEncoding=$__centralN2Utf8;$OutputEncoding=$__centralN2Utf8;"
    @staticmethod
    def resolve_host(host):
        try:return socket.gethostbyname(host)
        except OSError:return None
    @staticmethod
    def is_local(host):return HostIdentity.is_local(host)
    def select_transport(self,host,*,refresh=False):return self.transport_manager.select(host,refresh=refresh).name
    def invalidate_transport(self,host):self.transport_manager.invalidate(host)
    def ping(self,host):
        if self.is_local(host):return CommandResult(True,"local",host,stdout="OK",transport="local")
        return self._run_local(["ping","-n","1","-w","1200",host],host=host,action="ping")
    def test_admin_share(self,host):
        if self.is_local(host):return CommandResult(True,"local",host,stdout="OK",transport="local")
        return self._run_local(["cmd.exe","/d","/c",f"dir \\\\{host}\\admin$ >nul 2>&1"],host=host,action="admin_share")
    def test_winrm(self,host):return self.winrm_transport.test(host)
    @staticmethod
    def _is_transport_failure(r):
        text=(r.stderr or "").lower();return not r.success and any(x in text for x in ("psremotingtransportexception","cannotconnect","pssessionstatebroken","ws-management","the client cannot connect","o cliente não conseguiu se conectar"))
    def execute_powershell(self,host,script,*,timeout=None):
        t=self.transport_manager.select(host);r=t.execute_powershell(host,script,timeout=timeout)
        if t.name=="winrm" and self._is_transport_failure(r) and self.psexec_transport.available():
            self.transport_manager.invalidate(host);f=self.psexec_transport.execute_powershell(host,script,timeout=timeout);f.metadata["fallback_from"]="winrm";return f
        return r
    def execute_powershell_json(self,host,script,*,timeout=None):
        r=self.execute_powershell(host,f"$r=& {{ {script} }};$r|ConvertTo-Json -Depth 8 -Compress",timeout=timeout)
        if r.success and r.stdout.strip():
            try:r.data=json.loads(r.stdout.strip())
            except json.JSONDecodeError:r.metadata["json_parse_error"]=True
        return r
    def execute_psexec(self,host,executable,args:Iterable[str]=(),*,system=False,timeout=None,output_encoding=None):return self.psexec_transport.execute_raw(host,executable,list(args),system=system,timeout=timeout,output_encoding=output_encoding)
    def execute_cmd(self,host,command,*,timeout=None):
        t=self.transport_manager.select(host);r=t.execute_cmd(host,command,timeout=timeout)
        if t.name=="winrm" and self._is_transport_failure(r) and self.psexec_transport.available():
            self.transport_manager.invalidate(host);f=self.psexec_transport.execute_cmd(host,command,timeout=timeout);f.metadata["fallback_from"]="winrm";return f
        return r
    def execute_remote_powershell_with_fallback(self,host,script,*,timeout=None):return self.execute_powershell(host,script,timeout=timeout)
    def execute_cmd_stream(self,host,command,*,timeout=None,on_line:Callable[[str],None]|None=None):
        t=self.transport_manager.select(host)
        if t.name=="local":cmd=["cmd.exe","/d","/c",command];enc=None
        elif t.name=="winrm":
            safe=host.replace("'","''");escaped=command.replace("'","''");payload=self._powershell_utf8_prefix()+f"$ErrorActionPreference='Stop';Invoke-Command -ComputerName '{safe}' -ScriptBlock {{ cmd.exe /d /c '{escaped}' }}"
            cmd=["powershell.exe","-NoProfile","-NonInteractive","-Command",payload];enc="utf-8"
        else:cmd=[str(self.psexec_path),"-accepteula","-nobanner",f"\\\\{host}","cmd.exe","/d","/c",command];enc=None
        r=self._run_local_streaming(cmd,host=host,action="cmd_stream",timeout=timeout,output_encoding=enc,on_line=on_line);r.transport=t.name;return r
    def _run_powershell_local(self,script,*,host,action,timeout=None):return self._run_local(["powershell.exe","-NoProfile","-NonInteractive","-Command",self._powershell_utf8_prefix()+script],host=host,action=action,timeout=timeout,output_encoding="utf-8")
    def _run_local_streaming(
        self,
        cmd,
        *,
        host,
        action,
        timeout=None,
        output_encoding=None,
        on_line=None,
    ):
        started = time.perf_counter()
        printable = subprocess.list2cmdline(cmd)
        lines: list[str] = []
        process = None

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

            threading.Thread(
                target=reader,
                name="central-n2-stream-reader",
                daemon=True,
            ).start()

            effective_timeout = timeout or self.timeout
            deadline = time.monotonic() + effective_timeout

            while True:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    process.kill()
                    process.wait(timeout=5)
                    raise subprocess.TimeoutExpired(cmd, effective_timeout)

                try:
                    raw = events.get(timeout=min(0.1, remaining))
                except queue.Empty:
                    if process.poll() is not None and events.empty():
                        break
                    continue

                if raw is None:
                    break

                text = self._decode_output(
                    raw,
                    preferred=output_encoding,
                ).rstrip()
                lines.append(text)
                if on_line:
                    on_line(text)

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
    def _run_local(self,cmd,*,host,action,timeout=None,output_encoding=None):
        started=time.perf_counter();printable=subprocess.list2cmdline(cmd)
        try:
            p=subprocess.run(cmd,capture_output=True,text=False,timeout=timeout or self.timeout,shell=False);r=CommandResult(p.returncode==0,printable,host,stdout=self._decode_output(p.stdout,preferred=output_encoding).strip(),stderr=self._decode_output(p.stderr,preferred=output_encoding).strip(),return_code=p.returncode,duration_ms=int((time.perf_counter()-started)*1000))
        except subprocess.TimeoutExpired:r=CommandResult.failure(host,printable,f"Timeout após {timeout or self.timeout}s",return_code=124);r.duration_ms=int((time.perf_counter()-started)*1000)
        except OSError as exc:r=CommandResult.failure(host,printable,str(exc),return_code=127);r.duration_ms=int((time.perf_counter()-started)*1000)
        if self.logger:self.logger.log_result(action,r)
        return r
