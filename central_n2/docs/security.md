# Segurança — Central N2 Workstation 5.2.0

## Modelo de confiança

A Central é ferramenta administrativa privilegiada para operação autorizada em estações Windows.

Pressupostos:

- operador autorizado;
- estação administrativa confiável;
- rede/hosts dentro do escopo permitido;
- política corporativa respeitada;
- PsExec homologado quando utilizado.

## Princípio de menor superfície

A Central de Execuções não expõe:

- shell PowerShell livre;
- shell CMD livre;
- editor genérico de Registro;
- instalador arbitrário informado em runtime;
- PFX/chave privada;
- exclusão recursiva genérica.

Poder operacional deve entrar como ação tipada, testada e auditável.

## Policy antes da execução

Toda ação passa por:

- sessão READY;
- transporte permitido;
- privilege;
- capabilities;
- custom preconditions.

Bloqueio ocorre antes da confirmação e antes do handler.

## Privilege

Níveis suportados:

- ANY;
- ADMIN;
- SYSTEM;
- USER_CONTEXT.

Winget, por exemplo, pode exigir contexto de usuário administrativo e ser bloqueado sob SYSTEM/PsExec.

## Retry

`RetryPolicy` é parte do contrato.

Padrão: `PRE_EXECUTION_ONLY`.

`SAFE_TRANSIENT` só é válido para ação idempotente.

Resultado indeterminado não sofre repetição automática cega.

## Disconnect esperado

Ações que derrubam a própria conectividade devem declarar DisconnectMode.

TEMPORARY exige recovery + postcheck.

TERMINAL representa efeito final esperado, como shutdown.

## Rollback

Rollback só é oferecido quando:

- before probe captura estado suficiente;
- existe operação tecnicamente inversa;
- after probe consegue validar restauração.

A Central não simula rollback para ações irreversíveis.

## Allowlist corporativa

`packages`, `certificates` e `registry_actions` são validados no bootstrap.

Regras importantes:

- pacote: somente tipo homologado;
- certificado: somente .cer/.crt público e stores permitidos;
- Registro: somente HKLM e tipos permitidos.

Entrada inválida é removida do catálogo efetivo.

## Segredos

Nunca versione:

- senhas;
- tokens;
- API keys;
- Authorization headers;
- certificados privados;
- credenciais de domínio;
- caminhos/URLs internos sensíveis.

Use `settings.local.json`.

## Auditoria e redaction

`ExecutionRecord.audit_payload()` e SQLite aplicam redaction.

Parâmetros marcados `sensitive=True` são substituídos por `***`.

Campos comuns como password/token/secret também passam pelo redactor central.

## WinRM e PsExec

Evite:

- `TrustedHosts=*`;
- desligar Firewall/Defender para “fazer funcionar”;
- abrir listener fora da política;
- assumir que ADMIN$ implica PsExec funcional.

PsExec pode executar em contexto diferente do usuário interativo; policy/capabilities devem refletir isso.

## Concorrência

JobManager serializa mutações por host.

Isso evita sobreposição de operações como DISM, cleanup, alteração de rede e reboot na mesma estação.

## Persistência

SQLite e relatórios podem conter infraestrutura. Proteja com ACL e retenção.

Schema é migrado automaticamente e passa por quick_check.

## Repositório público

Nunca commite:

- logs reais;
- banco SQLite;
- inventários;
- nomes de usuários;
- dumps;
- relatórios;
- settings.local.json.

## Incidente

1. interrompa a operação;
2. preserve logs/banco;
3. identifique correlation_id;
4. identifique operador, host e action key;
5. revise execution history e jobs;
6. revise credenciais;
7. acione segurança corporativa.
