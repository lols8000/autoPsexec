from __future__ import annotations

import json
from pathlib import Path

from .models import ControlledActionSpec, EndpointSpec


def load_campaign(path: Path) -> tuple[str, list[EndpointSpec]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("A campanha deve ser um objeto JSON.")

    name = str(payload.get("name") or "central-n2-field-validation")
    raw_endpoints = payload.get("endpoints")
    if not isinstance(raw_endpoints, list):
        raise ValueError("'endpoints' deve ser uma lista.")

    endpoints: list[EndpointSpec] = []
    aliases: set[str] = set()
    for raw in raw_endpoints:
        if not isinstance(raw, dict):
            raise ValueError("Cada endpoint deve ser um objeto.")
        alias = str(raw.get("alias") or "").strip()
        target = str(raw.get("target") or "").strip()
        if not alias or not target:
            raise ValueError("Endpoint exige alias e target.")
        if alias.casefold() in aliases:
            raise ValueError(f"Alias duplicado: {alias}")
        aliases.add(alias.casefold())

        raw_actions = raw.get("actions", [])
        if not isinstance(raw_actions, list):
            raise ValueError(
                f"Endpoint {alias}: 'actions' deve ser uma lista."
            )

        actions = tuple(
            ControlledActionSpec(
                key=str(item.get("key") or "").strip(),
                parameters=dict(item.get("parameters") or {}),
                rollback_after=bool(
                    item.get("rollback_after", False)
                ),
            )
            for item in raw_actions
            if isinstance(item, dict)
        )
        endpoints.append(
            EndpointSpec(
                alias=alias,
                target=target,
                enabled=bool(raw.get("enabled", True)),
                expected_state=(
                    str(raw["expected_state"])
                    if raw.get("expected_state")
                    else None
                ),
                expected_transport=(
                    str(raw["expected_transport"])
                    if raw.get("expected_transport")
                    else None
                ),
                required_capabilities=tuple(
                    str(item)
                    for item in raw.get(
                        "required_capabilities",
                        [],
                    )
                ),
                roles=tuple(
                    str(item)
                    for item in raw.get("roles", [])
                ),
                actions=actions,
            )
        )

    return name, endpoints
