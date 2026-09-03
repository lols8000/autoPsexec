from __future__ import annotations

import time
from dataclasses import dataclass

from core.host_identity import HostIdentity
from core.result import CommandResult
from core.retry import RetryPolicy
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
        retry_policy: RetryPolicy | None = None,
    ) -> None:
        self.local = local
        self.winrm = winrm
        self.psexec = psexec
        self.cache_ttl_seconds = max(0.0, float(cache_ttl_seconds))
        self.retry_policy = retry_policy or RetryPolicy()
        self._cache: dict[str, CachedTransport] = {}

    def invalidate(self, host: str) -> None:
        self._cache.pop(host.lower(), None)

    def select(
        self,
        host: str,
        *,
        refresh: bool = False,
        winrm_result: CommandResult | None = None,
    ) -> Transport:
        key = host.lower()
        now = time.monotonic()
        cached = self._cache.get(key)
        if cached and cached.expires_at > now and not refresh:
            return cached.transport

        if HostIdentity.is_local(host):
            selected = self.local
        else:
            probe = winrm_result or self.retry_policy.run(
                lambda: self.winrm.test(host)
            )
            if probe.success:
                selected = self.winrm
            elif self.psexec.available():
                selected = self.psexec
            else:
                # Mantém WinRM como transporte nominal para que o erro
                # diagnóstico original seja preservado ao operador.
                selected = self.winrm

        self._cache[key] = CachedTransport(
            selected,
            now + self.cache_ttl_seconds,
        )
        return selected
