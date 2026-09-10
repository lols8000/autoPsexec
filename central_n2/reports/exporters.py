from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

_SAFE_STEM_RE = re.compile(r"[^A-Za-z0-9._-]+")


class ReportExporter:
    def __init__(self, directory: str | Path) -> None:
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _safe_stem(value: str) -> str:
        stem = _SAFE_STEM_RE.sub("_", value).strip("._")
        if not stem:
            raise ValueError("Nome de relatório inválido.")
        return stem[:180]

    @staticmethod
    def _format_value(value: Any) -> str:
        if value is None or value == "":
            return "-"
        if isinstance(value, (dict, list, tuple)):
            return json.dumps(
                value,
                ensure_ascii=False,
                indent=2,
                default=str,
            )
        return str(value)

    def export(
        self,
        report: dict[str, Any],
        *,
        fmt: str = "markdown",
        stem: str = "report",
    ) -> Path:
        normalized = fmt.casefold()
        safe_stem = self._safe_stem(stem)

        if normalized == "json":
            path = self.directory / f"{safe_stem}.json"
            path.write_text(
                json.dumps(
                    report,
                    ensure_ascii=False,
                    indent=2,
                    default=str,
                ),
                encoding="utf-8",
            )
            return path

        if normalized in {"txt", "text"}:
            path = self.directory / f"{safe_stem}.txt"
            path.write_text(
                self._text(report),
                encoding="utf-8",
            )
            return path

        if normalized not in {"markdown", "md"}:
            raise ValueError(
                f"Formato de relatório não suportado: {fmt}"
            )

        path = self.directory / f"{safe_stem}.md"
        path.write_text(
            self._markdown(report),
            encoding="utf-8",
        )
        return path

    @classmethod
    def _text(cls, report: dict[str, Any]) -> str:
        actions = (
            "\n".join(
                f"- {item}"
                for item in report.get("actions", [])
            )
            or "-"
        )
        validation = cls._format_value(
            report.get("validation")
        )
        return (
            "ATENDIMENTO N2\n\n"
            f"Correlation ID: "
            f"{report.get('correlation_id') or '-'}\n"
            f"Estação: {report.get('host') or '-'}\n"
            f"Usuário: {report.get('user') or '-'}\n"
            f"Problema: {report.get('problem') or '-'}\n\n"
            f"Diagnóstico:\n"
            f"{report.get('diagnosis') or '-'}\n\n"
            f"Ações:\n{actions}\n\n"
            f"Validação:\n{validation}\n\n"
            f"Resultado: {report.get('result') or '-'}\n"
        )

    @classmethod
    def _markdown(cls, report: dict[str, Any]) -> str:
        actions = (
            "\n".join(
                f"- {item}"
                for item in report.get("actions", [])
            )
            or "- Nenhuma ação registrada"
        )
        validation = cls._format_value(
            report.get("validation")
        )
        return (
            "# Atendimento N2\n\n"
            f"**Correlation ID:** "
            f"{report.get('correlation_id') or '-'}  \n"
            f"**Estação:** {report.get('host') or '-'}  \n"
            f"**Usuário:** {report.get('user') or '-'}  \n"
            f"**Gerado em:** {report.get('generated_at') or '-'}\n\n"
            f"## Problema\n"
            f"{report.get('problem') or '-'}\n\n"
            f"## Diagnóstico\n"
            f"{report.get('diagnosis') or '-'}\n\n"
            f"## Ações\n{actions}\n\n"
            f"## Validação\n"
            f"{validation}\n\n"
            f"## Resultado\n"
            f"{report.get('result') or '-'}\n"
        )
