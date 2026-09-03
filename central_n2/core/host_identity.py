from __future__ import annotations

import ipaddress
import os
import socket
from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class HostDescriptor:
    requested: str
    normalized: str
    is_local: bool
    resolved_addresses: tuple[str, ...]


class HostIdentity:
    @staticmethod
    def _normalize(value: str) -> str:
        return value.strip().strip("[]").rstrip(".").lower()

    @classmethod
    def local_names(cls) -> set[str]:
        values = {"localhost", ".", "127.0.0.1", "::1"}
        for candidate in (
            socket.gethostname(),
            socket.getfqdn(),
            os.environ.get("COMPUTERNAME"),
        ):
            if candidate:
                values.add(cls._normalize(candidate))
                values.add(cls._normalize(candidate.split(".", 1)[0]))
        return values

    @classmethod
    def local_addresses(cls) -> set[str]:
        addresses = {"127.0.0.1", "::1"}
        for name in cls.local_names():
            if name in {".", "localhost"}:
                continue
            try:
                for item in socket.getaddrinfo(name, None):
                    addresses.add(item[4][0].split("%", 1)[0])
            except OSError:
                continue
        return addresses

    @classmethod
    def resolve_all(cls, host: str) -> tuple[str, ...]:
        values: set[str] = set()
        try:
            for item in socket.getaddrinfo(host, None):
                values.add(item[4][0].split("%", 1)[0])
        except OSError:
            pass
        return tuple(sorted(values))

    @classmethod
    def is_local(cls, host: str) -> bool:
        normalized = cls._normalize(host)
        if normalized in cls.local_names():
            return True
        try:
            ip = ipaddress.ip_address(normalized.split("%", 1)[0])
            if ip.is_loopback:
                return True
            return str(ip) in cls.local_addresses()
        except ValueError:
            pass
        resolved = cls.resolve_all(host)
        return bool(set(resolved) & cls.local_addresses())

    @classmethod
    def describe(cls, host: str) -> HostDescriptor:
        return HostDescriptor(
            requested=host,
            normalized=cls._normalize(host),
            is_local=cls.is_local(host),
            resolved_addresses=cls.resolve_all(host),
        )
