from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Callable

from core.result import CommandResult

RunLocal = Callable[..., CommandResult]
Utf8Prefix = Callable[[], str]


class Transport(ABC):
    name = "unknown"

    @abstractmethod
    def available(self) -> bool:
        raise NotImplementedError

    @abstractmethod
    def test(self, host: str) -> CommandResult:
        raise NotImplementedError

    @abstractmethod
    def execute_powershell(self, host: str, script: str, *, timeout: int | None = None) -> CommandResult:
        raise NotImplementedError

    @abstractmethod
    def execute_cmd(self, host: str, command: str, *, timeout: int | None = None) -> CommandResult:
        raise NotImplementedError
