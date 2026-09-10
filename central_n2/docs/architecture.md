# Arquitetura — Central N2 Workstation 5.1.0

## Visão

\`\`\`text
main.py
  ↓
ConfigLoader + AuditLogger
  ↓
RemoteExecutor
  ↓
ConsoleUIV5 / GUI
  ↓
JobManager + AttendanceContext + SessionManager
  ↓
Módulos / Playbooks / RemediationEngine
  ↓
LocalTransport | WinRMTransport | PsExecTransport
  ↓
Estação Windows
  ↓
SQLite / relatório / GLPI
\`\`\`

A regra de dependência é descendente: UI orquestra, módulos encapsulam domínio, executor decide transporte e transportes não conhecem a UI.

## Bootstrap

\`main.py\`:

1. processa argumentos;
2. valida Windows;
3. solicita UAC quando necessário;
4. carrega \`settings.json\` e \`settings.local.json\` uma única vez;
5. cria \`AuditLogger\`;
6. cria \`RemoteExecutor\`;
7. injeta a mesma configuração na interface selecionada.

A configuração efetiva não é relida independentemente pelos componentes principais.

## Estado do atendimento

\`AttendanceContext\` é a fonte única de verdade para:

- host lógico do atendimento;
- \`correlation_id\`;
- sessão;
- snapshot de saúde;
- diagnósticos;
- playbook;
- remediação;
- caminho do relatório.

Ao selecionar outro host, um novo contexto é criado; dados do atendimento anterior não são reutilizados.

## Sessão lógica

\`SessionManager\` não mantém uma \`PSSession\` permanente. Ele guarda:

- transporte selecionado;
- diagnóstico de conectividade;
- capabilities;
- estado de prontidão.

## Preflight

A conectividade é avaliada em camadas:

\`\`\`text
DNS
 ↓
ping + TCP 445 + TCP 5985 + TCP 5986
 ↓
WinRM autenticado / ADMIN$
 ↓
PsExec real
 ↓
estado final + transporte
\`\`\`

Estados:

| Estado | Significado |
| --- | --- |
| READY_LOCAL | alvo é a própria estação |
| READY_WINRM | WinRM autenticado e com Invoke-Command validado |
| READY_PSEXEC | WinRM indisponível e PsExec/ADMIN$ validados |
| DNS_FAILED | nome não pôde ser resolvido |
| AUTHENTICATION_FAILED | rede responde, mas autenticação/autorização falhou |
| NETWORK_UNREACHABLE | nenhum caminho administrativo conhecido respondeu |
| NO_USABLE_TRANSPORT | host alcançável, porém sem transporte validado |

### WinRM

O teste não se limita a \`Test-WSMan\`. A Central também executa um \`Invoke-Command\` mínimo e verifica um marcador conhecido. Assim, listener disponível não é confundido com sessão autenticada utilizável.

### PsExec

É fallback de primeira classe, não “último hack”. O teste executa comando remoto real. Dependências típicas: TCP 445, ADMIN$, privilégio administrativo e binário homologado.

## Semântica de fallback

O executor classifica a falha WinRM antes de decidir fallback.

### Consulta

\`\`\`text
READ_ONLY + falha de transporte
  → pode repetir via PsExec
\`\`\`

### Mutação

\`\`\`text
falha pré-execução comprovada
  → pode repetir via PsExec

falha durante/depois de possível execução
  → indeterminate=True
  → fallback_suppressed=True
  → operador valida estado antes de repetir
\`\`\`

A regra vale para PowerShell e CMD mutáveis.

## CommandResult

Contrato central:

- \`success\`;
- \`command\`;
- \`host\`;
- \`stdout\`;
- \`stderr\`;
- \`return_code\`;
- \`duration_ms\`;
- \`transport\`;
- \`data\`;
- \`metadata\`.

\`indeterminate\` é derivado de metadata e significa: a ação pode ter atingido o destino, mas a Central não conseguiu confirmar o resultado final.

## Scheduler

Há um \`JobManager\` compartilhado.

Classes:

- \`READ_ONLY\`;
- \`HEAVY_READ\`;
- \`LIGHT_WRITE\`;
- \`HEAVY_WRITE\`;
- \`DISRUPTIVE\`.

Operações que não são \`READ_ONLY\` são serializadas por host. Estados de job:

\`QUEUED → RUNNING → SUCCESS | FAILED | TIMEOUT | CANCELLED\`.

Timeout local não prova encerramento remoto. Um job marcado TIMEOUT não volta para SUCCESS se o worker terminar depois.

## Avaliação

\`core/evaluation.py\` centraliza os estados:

- \`PASS\`;
- \`FAIL\`;
- \`UNKNOWN\`;
- \`NOT_APPLICABLE\`.

UNKNOWN representa ausência de evidência. N/A representa controle não exigido. Apenas PASS/FAIL entram no denominador do compliance.

## Diagnóstico e playbooks

\`DiagnosticEngine\` produz fatos (\`Finding\`). \`CorrelationEngine\` produz diagnósticos (\`Diagnosis\`) com rationale e confiança.

Playbooks executam coletores orientados por sintoma; não são remediações automáticas.

## Remediação

\`RemediationEngine\` separa:

\`\`\`text
before probe
  ↓
ação
  ↓
after probe
  ↓
validador específico
  ↓
PASS / FAIL / UNKNOWN
\`\`\`

Se a execução for indeterminada, o validador deve preferir UNKNOWN quando não houver evidência suficiente.

## Persistência

SQLite:

- WAL;
- \`PRAGMA foreign_keys=ON\`;
- \`busy_timeout\`;
- migrations sequenciais;
- \`PRAGMA user_version\`;
- \`quick_check\` na inicialização;
- retenção configurável.

Schema atual: **2**.

Tabelas operacionais: \`hosts\`, \`snapshots\`, \`jobs\`, \`findings\`, \`remediations\`, \`reports\`.

## Auditoria

\`AuditLogger\` é thread-safe. Por padrão grava metadados compactos; payloads verbosos são opt-in. Redaction cobre senha, token, API key, Authorization, Bearer e credenciais.

## Distribuição

- PyInstaller onedir;
- UPX desativado;
- metadata de versão Windows;
- Inno Setup;
- release com SHA256SUMS;
- assinatura Authenticode opcional quando o CI recebe certificado.

## Regra de evolução

Nova funcionalidade deve respeitar:

\`\`\`text
entrada validada
+ ActionSpec
+ operação classificada
+ timeout
+ CommandResult
+ resultado estruturado quando possível
+ logging seguro
+ teste
+ documentação
\`\`\`
