from __future__ import annotations
from .base import PlaybookSpec,PlaybookStep
def _s(k,l): return PlaybookStep(k,l)
def builtin_playbooks():
    specs=(
      PlaybookSpec("slow","Computador lento",(_s("health","Saúde geral"),_s("performance","Performance"),_s("disk","Disco"),_s("processes","Processos"),_s("startup","Inicialização"))),
      PlaybookSpec("network","Sem internet / rede",(_s("network","IP/DNS/Gateway"),_s("adapters","Adaptadores"),_s("connections","Conexões TCP"),_s("proxy","Proxy"))),
      PlaybookSpec("printer","Não imprime",(_s("printers","Impressoras"),_s("print_queue","Fila"),_s("services","Serviços"))),
      PlaybookSpec("domain","Problema de domínio / GPO",(_s("domain","Domínio"),_s("gpresult","GPResult"))),
      PlaybookSpec("update","Windows Update",(_s("updates","Windows Update"),_s("health","Saúde"))),
      PlaybookSpec("crash","Aplicativo fechando",(_s("app_crashes","Crashes"),_s("processes","Processos"))),
      PlaybookSpec("bsod","Tela azul / BSOD",(_s("bsod","BugChecks/dumps"),_s("devices","Dispositivos"))),
      PlaybookSpec("disk","Disco cheio",(_s("disk","Espaço"),_s("profiles","Perfis"),_s("cleanup_estimate","Estimativa limpeza"))),
      PlaybookSpec("glpi","GLPI Agent",(_s("glpi","Status"),_s("glpi_log","Log recente"))),
    )
    return {x.key:x for x in specs}
