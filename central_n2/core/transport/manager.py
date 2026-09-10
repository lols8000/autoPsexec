from __future__ import annotations

import threading
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
    """Seleciona e cacheia o transporte realmente utilizável por host."""

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
        self._guard = threading.RLock()

    def invalidate(self, host: str) -> None:
        with self._guard:
            self._cache.pop(host.casefold(), None)

    def select(
        self,
        host: str,
        *,
        refresh: bool = False,
        winrm_result: CommandResult | None = None,
        psexec_result: CommandResult | None = None,
    ) -> Transport:
        key = host.casefold()
        now = time.monotonic()

        with self._guard:
            cached = self._cache.get(key)
            if cached and cached.expires_at > now and not refresh:
                return cached.transport

        if HostIdentity.is_local(host):
            selected = self.local
        else:
            winrm_probe = winrm_result or self.retry_policy.run(
                lambda: self.winrm.test(host)
            )
            if winrm_probe.success:
                selected = self.winrm
            elif self.psexec.available():
                psexec_probe = psexec_result or self.retry_policy.run(
                    lambda: self.psexec.test(host)
                )
                selected = self.psexec if psexec_probe.success else self.winrm
            else:
                # Sem fallback utilizável, preserva WinRM para que o erro de
                # conectividade/autenticação continue visível ao operador.
                selected = self.winrm

        with self._guard:
            self._cache[key] = CachedTransport(
                selected,
                now + self.cache_ttl_seconds,
            )
        return selected
