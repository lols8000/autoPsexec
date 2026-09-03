from __future__ import annotations

import time
from dataclasses import dataclass

from core.host_identity import HostIdentity
from .base import Transport


@dataclass(slots=True)
class CachedTransport:
    transport: Transport
    expires_at: float


class TransportManager:
    def __init__(
        self,
        local: Transport,
        winrm: Transport,
        psexec: Transport,
        *,
        cache_ttl_seconds: float = 120.0,
    ) -> None:
        self.local = local
        self.winrm = winrm
        self.psexec = psexec
        self.cache_ttl_seconds = max(0.0, float(cache_ttl_seconds))
        self._cache: dict[str, CachedTransport] = {}

    def invalidate(self, host: str) -> None:
        self._cache.pop(host.lower(), None)

    def select(self, host: str, *, refresh: bool = False) -> Transport:
        key = host.lower()
        now = time.monotonic()
        cached = self._cache.get(key)
        if cached and cached.expires_at > now and not refresh:
            return cached.transport

        if HostIdentity.is_local(host):
            selected = self.local
        else:
            winrm_result = self.winrm.test(host)
            if winrm_result.success:
                selected = self.winrm
            elif self.psexec.available():
                selected = self.psexec
            else:
                selected = self.winrm

        self._cache[key] = CachedTransport(selected, now + self.cache_ttl_seconds)
        return selected
