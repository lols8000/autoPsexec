from __future__ import annotations

from .catalogs import (
    certificates,
    devices,
    disk,
    domain,
    energy,
    files,
    glpi,
    network,
    packages,
    printers,
    processes,
    registry,
    security,
    services,
    software,
    updates,
    users,
    windows,
)
from .catalogs.common import ExecutionDependencies
from .registry import ActionRegistry


_CATALOGS = (
    processes,
    services,
    software,
    network,
    printers,
    devices,
    windows,
    updates,
    domain,
    users,
    disk,
    glpi,
    security,
    energy,
    packages,
    certificates,
    registry,
    files,
)


def build_execution_registry(
    deps: ExecutionDependencies,
) -> ActionRegistry:
    registry = ActionRegistry()
    for catalog in _CATALOGS:
        catalog.register(registry, deps)
    return registry


__all__ = [
    "ExecutionDependencies",
    "build_execution_registry",
]
