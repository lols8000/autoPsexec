from __future__ import annotations
import json
from pathlib import Path
class ReportExporter:
    def __init__(self,directory): self.directory=Path(directory); self.directory.mkdir(parents=True,exist_ok=True)
    def export(self,report,*,fmt="markdown",stem="report"):
        fmt=fmt.lower()
        if fmt=="json":
            p=self.directory/f"{stem}.json"; p.write_text(json.dumps(report,ensure_ascii=False,indent=2,default=str),encoding="utf-8"); return p
        if fmt in {"txt","text"}:
            p=self.directory/f"{stem}.txt"; p.write_text(self._text(report),encoding="utf-8"); return p
        p=self.directory/f"{stem}.md"; p.write_text(self._markdown(report),encoding="utf-8"); return p
    @staticmethod
    def _text(r):
        acts="\n".join(f"- {x}" for x in r.get("actions",[])) or "-"
        return f"ATENDIMENTO N2\n\nCorrelation ID: {r.get('correlation_id') or '-'}\nEstação: {r.get('host')}\nUsuário: {r.get('user') or '-'}\nProblema: {r.get('problem') or '-'}\n\nDiagnóstico:\n{r.get('diagnosis') or '-'}\n\nAções:\n{acts}\n\nValidação:\n{r.get('validation') or '-'}\n\nResultado: {r.get('result') or '-'}\n"
    @classmethod
    def _markdown(cls,r):
        acts="\n".join(f"- {x}" for x in r.get("actions",[])) or "- Nenhuma ação registrada"
        return f"# Atendimento N2\n\n**Correlation ID:** {r.get('correlation_id') or '-'}  \n**Estação:** {r.get('host')}  \n**Usuário:** {r.get('user') or '-'}  \n**Gerado em:** {r.get('generated_at')}\n\n## Problema\n{r.get('problem') or '-'}\n\n## Diagnóstico\n{r.get('diagnosis') or '-'}\n\n## Ações\n{acts}\n\n## Validação\n{r.get('validation') or '-'}\n\n## Resultado\n{r.get('result') or '-'}\n"
