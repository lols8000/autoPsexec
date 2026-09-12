from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from .models import CampaignResult


def write_reports(
    campaign: CampaignResult,
    output_dir: Path,
) -> tuple[Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    json_path = output_dir / f"endpoint-validation-{stamp}.json"
    md_path = output_dir / f"endpoint-validation-{stamp}.md"

    json_path.write_text(
        json.dumps(
            campaign.public_dict(),
            indent=2,
            ensure_ascii=False,
            default=str,
        ),
        encoding="utf-8",
    )

    lines = [
        f"# Homologação de endpoints — {campaign.name}",
        "",
        f"Status geral: **{campaign.status.value}**",
        f"Início: {campaign.started_at}",
        f"Fim: {campaign.finished_at}",
        "",
    ]
    for endpoint in campaign.endpoints:
        lines.extend(
            [
                f"## {endpoint.alias}",
                "",
                f"- Status: **{endpoint.status.value}**",
                (
                    "- Target fingerprint: "
                    f"{endpoint.target_fingerprint}"
                ),
                (
                    "- Correlation ID: "
                    f"{endpoint.correlation_id}"
                ),
                "",
                "| Check | Estado | Mensagem |",
                "|---|---|---|",
            ]
        )
        for check in endpoint.checks:
            message = check.message.replace("|", "\\|").replace(
                "\n",
                " ",
            )
            lines.append(
                f"| {check.key} | **{check.state.value}** | "
                f"{message} |"
            )
        lines.append("")

    md_path.write_text("\n".join(lines), encoding="utf-8")
    return json_path, md_path
