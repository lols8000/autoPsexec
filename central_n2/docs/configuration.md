# Configuração — Central N2 Workstation v5

## Arquivos

Base pública:

~~~text
central_n2\config\settings.json
~~~

Override local:

~~~text
central_n2\config\settings.local.json
~~~

O ConfigLoader faz merge recursivo: chaves do arquivo local substituem apenas os valores correspondentes, preservando o restante da configuração pública.

## Exemplo atual

~~~json
{
  "timeout_seconds": 60,
  "psexec_path": "C:\\Windows\\System32\\PsExec.exe",
  "sysinternals_dir": "C:\\Sysinternals",
  "runtime": {
    "transport_cache_ttl_seconds": 120,
    "max_workers": 6,
    "max_batch_workers": 5,
    "retry_attempts": 2,
    "retry_base_delay_seconds": 0.5
  },
  "glpi_api": {
    "enabled": false,
    "base_url": "",
    "app_token": "",
    "user_token": ""
  },
  "persistence": {
    "enabled": true,
    "database": "data/central_n2.db",
    "snapshot_retention_days": 180
  },
  "updates": {
    "enabled": true,
    "repository": "lols8000/autoPsexec",
    "channel": "stable"
  }
}
~~~

## PsExec

psexec_path define o caminho preferencial. Se o caminho configurado não existir, o executor também tenta descobrir:

~~~text
PsExec.exe no PATH
C:\Windows\System32\PsExec.exe
C:\Sysinternals\PsExec.exe
~~~

Validação recomendada:

~~~powershell
Test-Path C:\Sysinternals\PsExec.exe
Get-Command PsExec.exe -ErrorAction SilentlyContinue
~~~

O diretório C:\Sysinternals é recomendado para manter a suíte separada do System32.

## Sysinternals

sysinternals_dir aponta para o diretório das ferramentas opcionais, como Autorunsc, ProcDump, Handle e Sigcheck.

## Runtime

- transport_cache_ttl_seconds: tempo de cache da seleção de transporte;
- max_workers: limite do JobManager;
- max_batch_workers: limite para lote;
- retry_attempts: tentativas de preflight/transporte para falhas transitórias;
- retry_base_delay_seconds: backoff inicial.

Retry não deve ser usado para repetir cegamente remediações destrutivas.

## GLPI Agent

A seção glpi configura instalador, caminho remoto e nomes possíveis do serviço.

Valores internos devem ficar em settings.local.json.

## GLPI API

A API fica desabilitada por padrão.

Exemplo local:

~~~json
{
  "glpi_api": {
    "enabled": true,
    "base_url": "https://glpi.exemplo/api",
    "app_token": "SEGREDO",
    "user_token": "SEGREDO"
  }
}
~~~

Nunca commite tokens.

## Software / Winget

A seção software define catálogo homologado por nome e winget_id.

Winget pode não funcionar no mesmo contexto em PsExec/SYSTEM.

## Compliance / Baseline

A configuração ativa um perfil e pode sobrescrever regras como:

- min_disk_free_percent;
- max_uptime_days;
- defender_required;
- firewall_required;
- glpi_required;
- pending_reboot_not_allowed.

Perfis versionados: DEFAULT, DESKTOP, NOTEBOOK e TI.

## Persistência

Quando habilitada, SQLite é criado no caminho configurado. Banco, WAL e relatórios operacionais não devem ser versionados.

## Updates

A seção updates define repositório e canal. A aplicação consulta releases e pode baixar artefatos, mas não se substitui silenciosamente.

## UI

heartbeat_seconds controla feedback visual. long_operation_timeout_seconds controla timeout padrão de operações longas.

Timeout local não garante encerramento de processo remoto já iniciado.

## O que nunca colocar no arquivo público

- senha;
- token;
- API key;
- credencial de domínio;
- segredo de proxy;
- chave privada;
- inventário real;
- dados pessoais;
- infraestrutura interna desnecessária.

## Checklist

~~~text
[ ] Python >= 3.10 ou pacote distribuído
[ ] PsExec homologado, se necessário
[ ] WinRM conforme política, se utilizado
[ ] settings.local.json fora do Git
[ ] Sysinternals homologado, se utilizado
[ ] GLPI configurado localmente
[ ] baseline aprovado
[ ] persistência protegida
[ ] testes em bancada
~~~
