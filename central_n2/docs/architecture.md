# Arquitetura — Central N2 Workstation 5.2.0

## Visão geral

```text
main.py
  ↓
ConfigLoader
  ↓
validate_execution_configuration
  ↓
AuditLogger + RemoteExecutor
  ↓
ConsoleUIV5
  ↓
AttendanceContext + SessionManager + JobManager
  ↓
Diagnóstico / Playbooks / Central de Execuções
  ↓
ExecutionRegistry → ExecutionPolicy → ExecutionEngine
  ↓
Módulos de domínio
  ↓
Local | WinRM | PsExec
  ↓
Estação Windows
  ↓
Postcheck / Recovery / Rollback
  ↓
SQLite v4 / Relatórios / GLPI
```

## Estado único do atendimento

`AttendanceContext` mantém:

- host;
- correlation_id;
- sessão;
- health snapshot;
- diagnoses;
- playbook;
- remediation;
- última execution;
- report path.

Trocar de host cria novo contexto e invalida evidência anterior.

## Sessão lógica

`SessionManager` não mantém PSSession permanente. Ele guarda:

- transporte selecionado;
- conectividade;
- capabilities;
- capability error;
- readiness.

## Preflight

```text
DNS
 ↓
ping + 445 + 5985 + 5986
 ↓
WinRM autenticado / ADMIN$
 ↓
PsExec real
 ↓
CapabilityDetector
 ↓
WorkstationSession
```

WinRM só fica READY após `Invoke-Command` autenticado.

## RemoteExecutor e fallback

Consultas podem repetir por fallback.

Mutações usam APIs específicas:

- `execute_mutating_powershell`;
- `execute_mutating_powershell_json`;
- `execute_mutating_cmd`.

Falha pré-execução comprovada pode cair para PsExec. Se a ação pode ter sido entregue, `CommandResult.indeterminate=True` e o fallback cego é suprimido.

## Central de Execuções

### Composition root

`execution/catalog.py` apenas compõe os catálogos.

Domínios ficam em `execution/catalogs/`:

- processes;
- services;
- software;
- network;
- printers;
- devices;
- windows;
- updates;
- domain;
- users;
- disk;
- glpi;
- security;
- energy;
- packages;
- certificates;
- registry;
- files.

### ExecutionAction

Contrato operacional:

- key estável;
- categoria e descrição;
- OperationClass;
- RiskLevel;
- timeout;
- confirmação;
- destructive/reboot/conectividade;
- parâmetros tipados;
- action_version;
- idempotent;
- RetryPolicy;
- retry_attempts / retry_delay_seconds;
- allowed_transports;
- required_capabilities;
- required_privilege;
- DisconnectMode;
- recovery timeout/delay;
- rollback_strategy;
- tags.

### Policy

`ExecutionPolicy` avalia antes da confirmação:

```text
session.ready
+ transport permitido
+ capabilities
+ privilege
+ custom preconditions
= LIBERADO | BLOQUEADO
```

Nenhuma ação bloqueada chega ao handler.

### RetryPolicy

- `NEVER`;
- `PRE_EXECUTION_ONLY`;
- `SAFE_TRANSIENT`.

Padrão: `PRE_EXECUTION_ONLY`.

O ExecutionEngine executa retry **seletivo e limitado**:

- `NEVER`: uma tentativa;
- `PRE_EXECUTION_ONLY`: nova tentativa apenas quando o resultado informa `transport_failure_kind=pre_execution`;
- `SAFE_TRANSIENT`: além da regra anterior, aceita falha explicitamente marcada como `retry_safe=true`.

`retry_attempts` e `retry_delay_seconds` fazem parte do contrato. Resultado `indeterminate` nunca é repetido automaticamente.

O fallback WinRM → PsExec continua sendo responsabilidade do RemoteExecutor; retry do engine não substitui nem enfraquece essa regra. `SAFE_TRANSIENT` é aceito apenas para ação idempotente.

### DisconnectMode

- `NONE`: fluxo normal;
- `TEMPORARY`: queda esperada, aguarda recovery e reabre sessão;
- `TERMINAL`: perda de conexão é consequência final esperada, como shutdown.

Em TEMPORARY, um comando indeterminado só pode terminar em PASS se a estação voltar e o postcheck comprovar o objetivo.

### Rollback

`BoundExecutionAction` pode possuir `rollback_handler`, `rollback_validator` e `rollback_preconditions`.

As preconditions do rollback são independentes das preconditions da ida. Isso evita bloquear uma reversão porque o estado esperado após a ação já não satisfaz a condição original.

Rollback recebe:

- parâmetros originais;
- evidência before;
- host.

O rollback é uma nova execução auditada ligada à execução original.

## Seletores

`SelectorKind` suporta:

- PROCESS;
- SERVICE;
- ADAPTER;
- PRINTER;
- PROFILE;
- DEVICE;
- SESSION.

A UI consulta o módulo correspondente e oferece seleção amigável.

## JobManager

Classes:

- READ_ONLY;
- HEAVY_READ;
- LIGHT_WRITE;
- HEAVY_WRITE;
- DISRUPTIVE.

Operações não READ_ONLY são serializadas por host.

## Validação

Fluxo:

```text
before probe
  ↓
handler
  ↓
recovery opcional
  ↓
after probe
  ↓
validator
  ↓
PASS | FAIL | UNKNOWN
```

UNKNOWN significa que o estado final não pôde ser comprovado.

## Persistência

SQLite schema **4**.

Tabelas principais:

- hosts;
- snapshots;
- jobs;
- findings;
- remediations;
- executions;
- reports.

`executions` registra operador, action_version, transporte, timestamps, duration, risco, parâmetros redigidos, validation state, correlation_id, rollback_of e is_rollback.

## Auditoria

`ExecutionRecord.audit_payload()` persiste apenas representação sanitizada:

- parâmetros sensíveis redigidos;
- policy checks;
- metadata operacional selecionada de retry/fallback;
- transporte e return code;
- before/after redigidos;
- validation;
- recovery;
- rollback availability.

O comando bruto não é necessário para a auditoria padrão. O redactor também percorre dataclasses, evitando que segredos dentro de `CommandResult` escapem durante a serialização.

## Distribuição

- PyInstaller onedir;
- metadata de versão Windows;
- Inno Setup;
- SHA256SUMS;
- Authenticode opcional;
- CI em Python 3.10/3.12/3.13.
