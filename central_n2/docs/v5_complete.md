# Central N2 Workstation 5.1.0 — arquitetura consolidada

A 5.1 transforma a evolução histórica do autoPsexec em uma plataforma de suporte N2 com contrato explícito de transporte, estado, segurança de mutação, diagnóstico, validação e rastreabilidade.

## Fluxo

\`\`\`text
Alvo
 ↓
HostIdentity
 ↓
ConnectivityDiagnostics
 ↓
TransportManager
 ↓
Local / WinRM / PsExec
 ↓
AttendanceContext + SessionManager
 ↓
JobManager
 ↓
Snapshot / Coleta
 ↓
Evaluation + Finding
 ↓
Correlation / Diagnosis
 ↓
Playbook
 ↓
Remediation
 ↓
Validação
 ↓
SQLite / Relatório / GLPI
\`\`\`

## O que diferencia a 5.1

### WinRM realmente validado

\`Test-WSMan\` sozinho não basta. O preflight exige execução de \`Invoke-Command\`.

### PsExec como transporte suportado

5985 bloqueada não encerra o atendimento se 445, ADMIN$ e PsExec estiverem válidos.

### Mutação sem dupla execução

Leituras podem repetir por fallback. Mutações só repetem quando a falha é comprovadamente anterior à execução.

Se a entrega for incerta:

\`\`\`text
INDETERMINADO → validar estado → decidir
\`\`\`

### Estado único

Dados do atendimento ficam em \`AttendanceContext\`, eliminando mirrors paralelos de sessão/diagnóstico/remediação/relatório.

### Scheduler único

\`JobManager\` controla concorrência, heartbeat e serialização por host.

### Avaliação explícita

PASS, FAIL, UNKNOWN e N/A têm significados distintos.

### Remediação validada

Ação concluída não é sinônimo de problema resolvido. Remediações possuem probes e validadores específicos.

### Persistência evolutiva

SQLite usa migrations e \`PRAGMA user_version\`, com schema atual 2.

### Distribuição endurecida

- SemVer;
- download temporário;
- validação de tamanho;
- SHA-256 quando publicado;
- rename atômico;
- UPX desativado;
- hashes de release;
- assinatura opcional;
- CI fixado e reprodutível.

## Playbooks

- lentidão;
- rede;
- impressão;
- domínio/GPO;
- Windows Update;
- crash;
- BSOD;
- disco cheio;
- GLPI Agent.

Playbook é coleta orientada, não automação de remediação.

## Remediações guiadas

- limpeza segura;
- reinício de Spooler;
- reset de Windows Update;
- GPUpdate /force.

Estados de validação: PASS, FAIL, UNKNOWN.

## Baselines

- DEFAULT;
- DESKTOP;
- NOTEBOOK;
- TI.

Controles opcionais não exigidos aparecem como N/A.

## Interfaces

Console é a interface operacional principal.

A GUI Tkinter usa o modelo compartilhado de jobs para não manter um scheduler paralelo de execução.

## Compatibilidade histórica

Arquivos antigos permanecem no repositório para referência, mas não definem a arquitetura da 5.1.

\`ConsoleUIV5\` herda \`ConsoleBase\`, não \`ConsoleUIV3\`.

## Definition of Done 5.1

Uma mudança está pronta quando:

\`\`\`text
contrato correto
+ segurança de execução
+ resultado observável
+ teste
+ lint/type/coverage
+ build
+ documentação
+ revisão de segredo
\`\`\`
