# Configuração — Central N2 Workstation 5.1.0

## Princípio

A configuração efetiva é composta por:

\`\`\`text
config/settings.json
        +
config/settings.local.json (opcional, não versionado)
        ↓
merge recursivo
        ↓
configuração carregada uma única vez no bootstrap
\`\`\`

Segredos e valores internos ficam somente em \`settings.local.json\`.

## Configuração pública atual

\`\`\`json
{
  "timeout_seconds": 60,
  "psexec_path": "C:\\Windows\\System32\\PsExec.exe",
  "sysinternals_dir": "C:\\Sysinternals",
  "runtime": {
    "transport_cache_ttl_seconds": 120,
    "max_workers": 6,
    "retry_attempts": 2,
    "retry_base_delay_seconds": 0.5
  },
  "glpi": {
    "installer_source": "",
    "remote_installer_path": "C:\\glpiagentinstall.vbs",
    "service_names": ["glpi-agent", "GLPI-Agent"]
  },
  "glpi_api": {
    "enabled": false,
    "base_url": "",
    "app_token": "",
    "user_token": ""
  },
  "software": {
    "chrome": {"name": "Google Chrome", "winget_id": "Google.Chrome"},
    "firefox": {"name": "Mozilla Firefox", "winget_id": "Mozilla.Firefox"},
    "7zip": {"name": "7-Zip", "winget_id": "7zip.7zip"},
    "vnc": {"name": "UltraVNC", "winget_id": "uvncbvba.UltraVnc"}
  },
  "compliance": {
    "profile": "DEFAULT",
    "overrides": {}
  },
  "persistence": {
    "enabled": true,
    "database": "data/central_n2.db",
    "snapshot_retention_days": 180
  },
  "updates": {
    "enabled": true,
    "repository": "lols8000/autoPsexec"
  },
  "ui": {
    "heartbeat_seconds": 0.2,
    "long_operation_timeout_seconds": 3600
  },
  "logging": {
    "verbose_payloads": false,
    "max_error_chars": 4000
  }
}
\`\`\`

## timeout_seconds

Timeout padrão do \`RemoteExecutor\` para operações que não definem valor mais específico.

Operações longas podem usar timeout próprio.

## psexec_path

Caminho preferencial para PsExec. Se não existir, a Central tenta descobrir:

- PsExec.exe no PATH;
- \`C:\Windows\System32\PsExec.exe\`;
- \`C:\Sysinternals\PsExec.exe\`.

Diretório recomendado: \`C:\Sysinternals\`.

## sysinternals_dir

Diretório das ferramentas opcionais Autorunsc, ProcDump, Handle e Sigcheck.

A Central não baixa Sysinternals automaticamente.

## runtime.transport_cache_ttl_seconds

Tempo de cache da escolha de transporte por host. Mudanças de conectividade podem ser forçadas com refresh pelo preflight.

## runtime.max_workers

Número máximo de workers do \`JobManager\` compartilhado.

Não representa “quantidade de remediações simultâneas por host”: operações não somente-leitura continuam serializadas por estação.

## runtime.retry_attempts / retry_base_delay_seconds

Retry é destinado a falhas transitórias de transporte/preflight. Não deve ser interpretado como repetição automática de remediação.

Mutações com resultado potencialmente entregue ao host são marcadas como indeterminadas e não recebem fallback cego.

## glpi

Configura o agente GLPI:

- \`installer_source\`;
- \`remote_installer_path\`;
- nomes possíveis de serviço.

Caminhos UNC internos devem ficar em \`settings.local.json\`.

## glpi_api

Desabilitada por padrão.

Exemplo local:

\`\`\`json
{
  "glpi_api": {
    "enabled": true,
    "base_url": "https://glpi.exemplo/apirest.php",
    "app_token": "PREENCHA_LOCALMENTE",
    "user_token": "PREENCHA_LOCALMENTE"
  }
}
\`\`\`

Nunca versione tokens.

## software

Catálogo permitido para operações Winget. Cada item possui nome amigável e \`winget_id\`.

Winget pode não existir no contexto SYSTEM/PsExec. Operações checam o exit code real do executável.

## compliance.profile

Perfis versionados:

- DEFAULT;
- DESKTOP;
- NOTEBOOK;
- TI.

A seleção interativa de perfil não é sobrescrita pelo perfil default da configuração.

## compliance.overrides

Overrides explícitos têm precedência sobre o arquivo do perfil.

Exemplo:

\`\`\`json
{
  "compliance": {
    "profile": "TI",
    "overrides": {
      "max_uptime_days": 7,
      "bitlocker_required": true
    }
  }
}
\`\`\`

Estados de avaliação:

- PASS;
- FAIL;
- UNKNOWN;
- NOT_APPLICABLE.

Controles não exigidos pelo baseline aparecem como N/A.

## persistence

\`enabled\`: ativa SQLite.

\`database\`: caminho do banco, relativo à raiz de \`central_n2\` quando não absoluto.

\`snapshot_retention_days\`: retenção aplicada na inicialização para snapshots, jobs, findings, remediações e relatórios.

O schema é migrado automaticamente por versão. Não edite \`PRAGMA user_version\` manualmente.

## updates

\`enabled\`: habilita/desabilita o menu de consulta/download.

\`repository\`: repositório GitHub usado para releases.

Não há atualização silenciosa do executável.

## ui

\`heartbeat_seconds\`: frequência de feedback do runner.

\`long_operation_timeout_seconds\`: timeout padrão para orquestrações longas da UI.

## logging

\`verbose_payloads=false\` é o padrão recomendado.

\`max_error_chars\` limita erro persistido.

Redaction continua sendo aplicada mesmo quando payload verboso é habilitado.

## Exemplo local recomendado

\`\`\`json
{
  "psexec_path": "C:\\Sysinternals\\PsExec.exe",
  "glpi": {
    "installer_source": "\\\\servidor\\share\\glpiagentinstall.vbs"
  },
  "glpi_api": {
    "enabled": false,
    "base_url": "https://glpi.exemplo/apirest.php",
    "app_token": "PREENCHA_LOCALMENTE",
    "user_token": "PREENCHA_LOCALMENTE"
  },
  "logging": {
    "verbose_payloads": false
  }
}
\`\`\`

## Política

Se uma chave pública não tiver efeito real em runtime, ela não deve permanecer em \`settings.json\`. Configuração não é documentação decorativa.
