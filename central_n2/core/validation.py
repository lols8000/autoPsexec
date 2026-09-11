from __future__ import annotations

import ipaddress
import re
from pathlib import PureWindowsPath

_HOST_RE = re.compile(r"^[A-Za-z0-9](?:[A-Za-z0-9._-]{0,251}[A-Za-z0-9])?$")
_SID_RE = re.compile(r"^S-1-(?:\d+-){1,14}\d+$", re.IGNORECASE)
_PROCESS_RE = re.compile(r"^[^\\/:*?\"<>|\r\n]{1,260}$")
_SAFE_NAME_RE = re.compile(r"^[^\r\n'\"]{1,260}$")
_INF_RE = re.compile(r"^oem\d+\.inf$", re.IGNORECASE)


def validate_host(value: str) -> str:
    host = value.strip()
    if not host:
        raise ValueError("Hostname/IP não pode ser vazio.")
    if host in {".", "localhost"}:
        return host
    try:
        ipaddress.ip_address(host.strip("[]").split("%", 1)[0])
        return host
    except ValueError:
        pass
    if not _HOST_RE.fullmatch(host):
        raise ValueError("Hostname contém caracteres inválidos.")
    return host


def validate_port(value: int | str) -> int:
    port = int(value)
    if not 1 <= port <= 65535:
        raise ValueError("Porta deve estar entre 1 e 65535.")
    return port


def validate_sid(value: str) -> str:
    sid = value.strip()
    if not _SID_RE.fullmatch(sid):
        raise ValueError("SID inválido.")
    return sid


def validate_process_name(value: str) -> str:
    name = value.strip()
    if not _PROCESS_RE.fullmatch(name):
        raise ValueError("Nome de processo inválido.")
    return name


def validate_windows_path(value: str) -> str:
    path = value.strip()
    if not path:
        raise ValueError("Caminho não pode ser vazio.")
    candidate = PureWindowsPath(path)
    if not candidate.drive:
        raise ValueError("Informe um caminho absoluto do Windows.")
    return str(candidate)


def quote_powershell_literal(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def quote_cmd_argument(value: str) -> str:
    # aspas internas são removidas para impedir quebra da linha de comando;
    # argumentos vindos da UI devem ser dados, não fragmentos de comando.
    return '"' + value.replace('"', "") + '"'



def validate_safe_name(value: str, *, label: str = "Nome") -> str:
    name = value.strip()
    if not _SAFE_NAME_RE.fullmatch(name):
        raise ValueError(f"{label} contém caracteres inválidos.")
    return name


def validate_unc_path(value: str) -> str:
    path = value.strip()
    if not path.startswith("\\\\"):
        raise ValueError("Informe um caminho UNC iniciado por \\\\.")
    if any(char in path for char in ("\r", "\n", '"')):
        raise ValueError("Caminho UNC contém caracteres inválidos.")
    return path


def validate_inf_name(value: str) -> str:
    name = value.strip()
    if not _INF_RE.fullmatch(name):
        raise ValueError("Informe um pacote INF no formato oemNN.inf.")
    return name
