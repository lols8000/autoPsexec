# Arquitetura — Central N2 Workstation v5

## Visão geral

A Central N2 é uma plataforma de troubleshooting de workstations Windows.

~~~text
UI
 ↓
JobManager / SessionManager
 ↓
Módulos
 ↓
RemoteExecutor
 ↓
Local / WinRM / PsExec
 ↓
Estação Windows
~~~

Depois da coleta, a v5 pode alimentar diagnóstico, correlação, playbook, remediação, persistência e relatório.

## Entry point

central_n2/main.py:

1. valida Windows;
2. localiza configuração;
3. solicita UAC;
4. carrega settings.json + settings.local.json;
5. cria AuditLogger;
6. cria RemoteExecutor;
7. inicia ConsoleUIV5 ou GUI.

## Interface

A interface principal é ConsoleUIV5, que herda menus funcionais da geração anterior e acrescenta:

- conectividade/capabilities;
- playbooks;
- histórico/diff;
- relatórios;
- jobs;
- atualização;
- GLPI API;
- remediações guiadas;
- baseline.

## Transportes

### Local

HostIdentity detecta localhost, hostname/FQDN local e endereços locais. Nesse caso não há WinRM nem PsExec.

### WinRM

Preferido quando o preflight indica disponibilidade.

### PsExec

Fallback quando WinRM não é utilizável e PsExec está disponível.

Dependências típicas:

- TCP 445/SMB;
- ADMIN$;
- privilégio administrativo;
- binário PsExec homologado.

O executor procura PsExec no PATH, System32 e C:\Sysinternals.

## Falha de WinRM não significa host offline

Exemplo:

~~~text
Ping ........... OK
TCP 445 ........ OK
ADMIN$ ......... OK
TCP 5985 ....... FAIL
PsExec ......... OK

Resultado: host administrável por PsExec.
~~~

A camada de conectividade deve separar DNS, ping, WinRM e ADMIN$.

## CommandResult

Todo resultado remoto usa CommandResult com:

- success;
- command;
- host;
- stdout;
- stderr;
- return_code;
- duration_ms;
- transport;
- data;
- metadata.

Dados estruturados devem preferir data em vez de parsing textual pela UI.

## JSON através de PsExec

PowerShell remoto pode ser envelopado por mensagens do PsExec/CLIXML.

RemoteExecutor._parse_json_output procura JSON válido mesmo quando existe ruído antes/depois.

PsExecTransport._clean_transport_noise remove apenas ruído de execução bem-sucedida. Erros reais permanecem.

## Apresentação

ConsoleUIV3 fornece renderização genérica de listas de objetos como tabela. Views específicas podem especializar a apresentação.

A view drivers:

- usa colunas fixas;
- converte booleano para SIM/NÃO;
- normaliza data;
- apresenta quantidade agrupada;
- mostra resumo.

## Jobs

JobManager usa ThreadPoolExecutor e classifica operações em:

- READ_ONLY;
- LIGHT_WRITE;
- HEAVY_WRITE;
- DISRUPTIVE.

Leituras podem concorrer. Escritas são serializadas por host.

Estados:

~~~text
QUEUED → RUNNING → SUCCESS / FAILED / TIMEOUT / CANCELLED
~~~

O estado TIMEOUT não é sobrescrito se o worker terminar depois.

## Retry

RetryPolicy é aplicado a falhas transitórias de transporte/preflight. Access Denied e falhas determinísticas não são repetidas indiscriminadamente.

## Sessão lógica

SessionManager guarda contexto por host:

- transporte;
- conectividade;
- capabilities.

Não é PSSession permanente.

## Diagnóstico

DiagnosticEngine transforma evidências em Finding. CorrelationEngine transforma combinações em Diagnosis com confiança e rationale.

## Playbooks

PlaybookRunner executa coletores somente leitura em sequência orientada por sintoma.

## Remediação

RemediationEngine:

~~~text
before → ação → after → diff
~~~

## Persistência

SQLite usa WAL e persiste snapshots, jobs, findings, remediações e relatórios.

## Auditoria

AuditLogger suporta correlation_id e sanitização de segredos.

## Configuração

ConfigLoader faz merge recursivo:

~~~text
settings.json
    +
settings.local.json
    ↓
configuração efetiva
~~~

## Distribuição

PyInstaller gera onedir. Inno Setup gera instalador. O CI Windows valida compileall, pytest, build portátil e instalador.

## Regra arquitetural

Dependência deve apontar para baixo:

~~~text
UI → orquestração → módulos → executor → transporte
~~~

O transporte não deve conhecer a UI.
