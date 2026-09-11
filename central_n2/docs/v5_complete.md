# Central N2 Workstation 5.2.0 — arquitetura consolidada

A 5.2 transforma a Central de uma plataforma predominantemente diagnóstica em uma plataforma operacional completa, mantendo o princípio de execução controlada.

## Fluxo consolidado

```text
Alvo
 ↓
HostIdentity
 ↓
ConnectivityDiagnostics
 ↓
Local / WinRM / PsExec
 ↓
SessionManager + Capabilities
 ↓
AttendanceContext
 ↓
Diagnóstico / Playbook
 ↓
ExecutionRegistry
 ↓
ExecutionPolicy
 ↓
Plano
 ↓
Confirmação
 ↓
ExecutionEngine
 ↓
Módulo de domínio
 ↓
Recovery opcional
 ↓
Postcheck
 ↓
PASS / FAIL / UNKNOWN
 ↓
SQLite v4
 ↓
Rollback quando suportado
```

## O que diferencia a 5.2

### Catálogo modular

Ações não ficam em um arquivo gigante. Cada domínio possui catálogo próprio e `catalog.py` apenas os compõe.

### Ação como contrato

A ação conhece risco, timeout, transportes, privilege, capabilities, retry, idempotência, disconnect e rollback.

### Policy antes da confirmação

A Central não pergunta “tem certeza?” para uma ação que já sabe ser incompatível. Ela bloqueia primeiro.

### Seleção orientada a objetos

Processos, serviços, adaptadores, impressoras, perfis, PnP e sessões podem ser escolhidos a partir de inventário do próprio host.

### Busca

Ações podem ser localizadas por texto e tags no menu 28.

### Expected disconnect

Restart de NIC, DHCP e reboot não são tratados como falha genérica. O workflow sabe que a conexão pode cair.

### Rollback honesto

Há rollback somente onde o estado anterior é conhecido e restaurável.

### Allowlist corporativa

Pacotes, certificados e Registro são tipados e validados no bootstrap.

### Auditoria rica

SQLite v4 registra execution records e rollback linkage com parâmetros redigidos.

## Estados

### Transporte

READY_LOCAL, READY_WINRM, READY_PSEXEC e estados de falha explícitos.

### Validação

PASS, FAIL, UNKNOWN.

### Policy

PASS, FAIL, WARN.

### Disconnect

NONE, TEMPORARY, TERMINAL.

### Retry

NEVER, PRE_EXECUTION_ONLY, SAFE_TRANSIENT.

## Compatibilidade

Scripts históricos continuam no repositório, mas não definem a arquitetura atual.

`ConsoleUIV5` herda `ConsoleBase`.

## Definition of Done 5.2

```text
ação modular
+ contrato completo
+ policy
+ capability/precondition
+ mutação segura
+ recovery quando necessário
+ postcheck
+ rollback quando real
+ auditoria
+ testes de falha
+ Ruff/mypy/coverage
+ build/installer
+ documentação
```
