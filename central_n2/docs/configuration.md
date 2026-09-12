# Configuração — Central N2 Workstation 5.2.0

## Composição

```text
config/settings.json
        +
config/settings.local.json
        ↓
merge recursivo
        ↓
validate_execution_configuration
        ↓
configuração efetiva sanitizada
```

Segredos e valores internos devem ficar somente em `settings.local.json`.

## Configuração pública

A configuração pública contém runtime, transporte, GLPI, software, compliance, persistência, updates, UI e logging.

Catálogos corporativos podem existir vazios na configuração pública e ser preenchidos localmente:

- `packages`;
- `certificates`;
- `registry_actions`.

## PsExec

`psexec_path` é o caminho preferencial.

Descoberta adicional:

- PATH;
- `C:\Windows\System32\PsExec.exe`;
- `C:\Sysinternals\PsExec.exe`.

## Runtime

### transport_cache_ttl_seconds

TTL do transporte selecionado.

### max_workers

Workers do JobManager compartilhado.

Mutações continuam serializadas por host.

### retry_attempts / retry_base_delay_seconds

Referem-se a transporte/preflight.

Não significam retry genérico de ExecutionAction.

## Software

`software` é catálogo Winget homologado.

Exemplo:

```json
{
  "software": {
    "chrome": {
      "name": "Google Chrome",
      "winget_id": "Google.Chrome"
    }
  }
}
```

A policy pode bloquear Winget quando o contexto remoto não é adequado, especialmente SYSTEM/PsExec.

## Pacotes corporativos

Exemplo local:

```json
{
  "packages": {
    "erp_client": {
      "name": "ERP Client",
      "source": "\\\\servidor\\software\\erp.msi",
      "type": "msi",
      "args": "/qn /norestart",
      "cleanup": true,
      "timeout_seconds": 1800
    }
  }
}
```

Tipos permitidos:

- msi;
- exe;
- cmd;
- bat.

O operador escolhe apenas uma chave homologada; não informa instalador arbitrário.

## Certificados

Exemplo:

```json
{
  "certificates": {
    "ca_corporativa": {
      "name": "CA Corporativa",
      "source": "\\\\servidor\\certs\\ca.cer",
      "store": "Root"
    }
  }
}
```

Extensões permitidas:

- .cer;
- .crt.

Stores permitidos:

- Root;
- CA;
- My;
- TrustedPeople.

PFX/chave privada não entra neste fluxo.

## Registry actions

Exemplo:

```json
{
  "registry_actions": {
    "produto_enabled": {
      "label": "Habilitar Produto",
      "path": "HKLM:\\SOFTWARE\\Empresa\\Produto",
      "name": "Enabled",
      "type": "DWord",
      "value": 1,
      "mode": "set"
    }
  }
}
```

Somente HKLM é aceito.

Modes:

- set;
- remove.

Tipos:

- String;
- ExpandString;
- DWord;
- QWord;
- MultiString;
- Binary.

## Validação no bootstrap

`validate_execution_configuration()` valida estrutura antes de inicializar os módulos.

Comportamento:

```text
entrada válida
  → permanece no catálogo

entrada inválida
  → removida da configuração efetiva
  → warning no console/log
  → Central continua disponível
```

A validação é estrutural. Um share temporariamente offline não deve invalidar o schema do catálogo.

## Execução / file_roots

Operações explícitas de arquivo são confinadas a raízes autorizadas.

Padrão:

```json
{
  "execution": {
    "file_roots": [
      "C:\\CentralN2",
      "C:\\Temp"
    ]
  }
}
```

`ensure directory`, `move/rename` e `remove file` recusam caminhos fora dessas raízes e caminhos que contenham `..`.

A validação ocorre no bootstrap. Se todas as raízes configuradas forem inválidas, a Central restaura os defaults seguros.

Essa configuração não libera shell nem exclusão recursiva genérica.

## Persistence

`persistence.enabled` ativa SQLite.

Schema atual: **4**.

`snapshot_retention_days` é aplicado a dados operacionais, incluindo executions.

## Logging

`verbose_payloads=false` é o recomendado.

Redaction é aplicada mesmo quando payload verboso está habilitado.

## Exemplo local mínimo

```json
{
  "psexec_path": "C:\\Sysinternals\\PsExec.exe",
  "packages": {},
  "certificates": {},
  "registry_actions": {},
  "glpi_api": {
    "enabled": false,
    "base_url": "",
    "app_token": "",
    "user_token": ""
  }
}
```

## Regra

Configuração deve representar comportamento real. Chave sem efeito runtime não deve permanecer como decoração.
