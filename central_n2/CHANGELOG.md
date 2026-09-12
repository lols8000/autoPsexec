# Changelog

## 5.2.0 — 2026-09-11

### Central de Execuções

- adicionado menu 28 com catálogo dinâmico, busca por texto/tags e atalhos contextuais;
- catálogo monolítico quebrado em 18 domínios sob `execution/catalogs/`;
- `ExecutionAction` passou a declarar risco, timeout, idempotência, retry, transportes, privilege, capabilities, disconnect mode, rollback e tags;
- registry valida invariantes de contrato antes de aceitar uma ação;
- parâmetros podem usar seletores de processos, serviços, adaptadores, impressoras, perfis, PnP e sessões;
- seletores de perfil/sessão passaram a usar inventários leves/estruturados, sem varredura recursiva ou parsing frágil de quser.

### Policy, recovery e segurança de mutação

- `ExecutionPolicy` avalia sessão, transporte, privilege, capabilities e preconditions antes da confirmação;
- ações incompatíveis são bloqueadas antes do handler;
- `RetryPolicy` formalizada com NEVER, PRE_EXECUTION_ONLY e SAFE_TRANSIENT;
- SAFE_TRANSIENT exige idempotência;
- ExecutionEngine aplica retry seletivo com retry_attempts/retry_delay_seconds apenas em falha pré-execução ou explicitamente retry-safe;
- resultado indeterminado nunca entra em retry automático;
- DisconnectMode NONE/TEMPORARY/TERMINAL;
- DHCP renew, restart de NIC e reboot podem usar recovery e postcheck após reconexão;
- resultado indeterminado continua preservado e só pode virar PASS após evidência pós-recovery suficiente.

### Rollback

- rollback real para Start/Stop e StartType de serviço;
- rollback de RegistryAction homologada para valor/ausência anterior;
- move/rename protegido pode retornar ao caminho original;
- rollback possui confirmação reforçada, validador e registro próprio;
- rollback_preconditions são independentes das preconditions da execução original;
- file.move bloqueia rollback quando a origem original voltou a existir, evitando overwrite;
- rollback PASS consome a disponibilidade de rollback do registro pai;
- ações irreversíveis não recebem rollback fictício;
- AttendanceContext mantém pilha LIFO de execuções reversíveis, preservando rollback mesmo após ações não reversíveis.

### Escopo de arquivos

- FileOperationsModule passou a exigir raízes autorizadas;
- defaults seguros: C:\CentralN2 e C:\Temp;
- traversal com .. e caminhos fora da allowlist são bloqueados antes do PowerShell;
- execution.file_roots é validado/sanitizado no bootstrap.

### Catálogos corporativos

- adicionadas allowlists de packages, certificates e registry_actions;
- configuração é validada/sanitizada no bootstrap;
- entradas inválidas são desabilitadas sem derrubar a aplicação;
- certificados privados/PFX e Registro fora de HKLM permanecem fora do fluxo.

### Auditoria e persistência

- SQLite promovido para schema 4;
- execution records armazenam operador, action_version, transporte, timestamps, duração, risco, parâmetros redigidos e correlation_id;
- rollbacks são ligados à execução original por rollback_of/is_rollback;
- `ExecutionRecord.audit_payload()` evita persistência de comando bruto e aplica redaction;
- auditoria persiste somente metadata operacional selecionada de retry/fallback;
- redactor central passou a percorrer dataclasses/Enums/coleções, protegendo `CommandResult` aninhado;
- persistência legada de execution também passa pelo redactor.

### Qualidade

- adicionada matriz reproduzível de homologação real de endpoints, com modo SAFE, execução controlada, relatórios sanitizados e smoke end-to-end em Windows real no CI;

- testes de policy, preconditions, disconnect/recovery, rollback, busca, allowlists, redaction, SQLite v4 e retry safety;
- pacote execution incluído em Ruff, mypy e coverage;
- documentação técnica revisada para 5.2.

## 5.1.0 — 2026-09-10

### Arquitetura

- Console v5 migrado para \`ConsoleBase\`, sem herança operacional de \`ConsoleUIV3\`.
- \`AttendanceContext\` tornou-se a fonte única de estado do atendimento para sessão, snapshot de saúde, diagnósticos, playbook, remediação, relatório e \`correlation_id\`.
- configuração carregada uma única vez no bootstrap e injetada no executor e nas interfaces;
- \`JobManager\` consolidado como scheduler da aplicação, com serialização por host para operações não somente-leitura;
- \`BatchRunner\` passou a poder utilizar o scheduler compartilhado;
- processos remotos rastreados passaram a usar os caminhos explícitos de mutação do executor.

### Transporte e segurança de execução

- o preflight WinRM só considera \`READY_WINRM\` após validar \`Test-WSMan\` e uma execução real de \`Invoke-Command\`;
- conectividade separa DNS, ping, TCP 445, TCP 5985/5986, ADMIN$, WinRM autenticado e PsExec;
- estados operacionais de conectividade: \`READY_LOCAL\`, \`READY_WINRM\`, \`READY_PSEXEC\`, \`DNS_FAILED\`, \`AUTHENTICATION_FAILED\`, \`NETWORK_UNREACHABLE\` e \`NO_USABLE_TRANSPORT\`;
- leituras podem fazer fallback WinRM → PsExec em falha de transporte;
- mutações só fazem fallback quando a falha é comprovadamente anterior à execução remota;
- quando uma mutação pode ter chegado ao destino e a confirmação se perde, o resultado é marcado como **indeterminado** e o fallback automático é bloqueado;
- resultados JSON de mutação que não puderem ser validados também são marcados como indeterminados;
- PsExec continua filtrando apenas ruído de transporte em execuções bem-sucedidas, preservando erros reais.

### Diagnóstico, compliance e remediação

- motor de avaliação unificado com estados \`PASS\`, \`FAIL\`, \`UNKNOWN\` e \`NOT_APPLICABLE\`;
- controles opcionais desabilitados pelo baseline agora aparecem explicitamente como N/A, sem entrar no denominador do compliance;
- métricas ausentes são \`UNKNOWN\`, não falso \`FAIL\`;
- remediações guiadas possuem validadores específicos para Spooler, limpeza segura, Windows Update e GPUpdate;
- reset de Windows Update registra restauração dos serviços originalmente ativos;
- limpeza segura retorna dados estruturados e não toca a Lixeira;
- Winget propaga \`$LASTEXITCODE\` diferente de zero como falha real.

### Persistência e observabilidade

- SQLite passou a usar migrations versionadas com \`PRAGMA user_version\`;
- schema atual: versão 2;
- \`correlation_id\` persiste em snapshots, jobs, findings, remediações e relatórios;
- inicialização executa \`PRAGMA quick_check\`;
- retenção remove registros antigos de snapshots, jobs, findings, remediações e relatórios;
- logger tornou-se thread-safe, compacto por padrão e continua aplicando redaction de segredos;
- payloads verbosos de auditoria permanecem opt-in.

### Atualização e distribuição

- versão promovida para 5.1.0 em runtime, VERSION, pyproject, metadata do EXE e instalador;
- updater passou a usar SemVer real;
- download de atualização é feito em arquivo temporário, com validação de tamanho e SHA-256 quando o release publica digest, seguido de substituição atômica;
- nomes de assets são validados contra path traversal;
- \`updates.enabled\` agora é respeitado em runtime;
- knobs sem efeito foram removidos da configuração pública;
- PyInstaller usa \`upx=False\` e metadata de versão Windows;
- release publica manifesto \`SHA256SUMS.txt\`;
- assinatura Authenticode é suportada opcionalmente quando certificado é fornecido via secrets do CI.

### CI / qualidade

- dependências de desenvolvimento/CI fixadas em \`requirements-dev.txt\`;
- matriz Windows com Python 3.10, 3.12 e 3.13;
- \`SyntaxWarning\` tratado como erro;
- Ruff como gate de correção;
- mypy aplicado aos contratos endurecidos;
- coverage mínimo como gate;
- build portátil PyInstaller e smoke do instalador Inno Setup;
- Inno Setup fixado em 6.7.1;
- GitHub Actions fixadas por SHA imutável;
- testes adicionados para bootstrap real da v5, estado único, migrations, logger compacto, SemVer, updater atômico, SHA-256, fallback seguro, WinRM autenticado, N/A de compliance e consistência de versão.

### Compatibilidade

- scripts históricos da raiz permanecem como referência;
- \`ConsoleUIV3\` permanece no repositório apenas como histórico/compatibilidade, não como base da UI operacional v5.

## 5.0.0 — 2026-09-03

### Added

- detecção de alvo local e transporte local;
- LocalTransport, WinRMTransport e PsExecTransport;
- TransportManager com cache e fallback;
- diagnóstico de conectividade e capabilities;
- SessionManager;
- JobManager e mutex por host;
- retry seletivo;
- progresso DISM;
- DiagnosticEngine e CorrelationEngine;
- nove playbooks;
- RemediationEngine com before/after;
- SQLite;
- diff de snapshots;
- baselines DEFAULT, DESKTOP, NOTEBOOK e TI;
- GLPI REST API opcional;
- correlation_id e sanitização;
- relatórios Markdown, JSON e TXT;
- console v5 e GUI Tkinter;
- atualização controlada;
- PyInstaller, Inno Setup e workflow de release.

### Security

- logger mascara senha, token, Authorization, API key, Bearer e credenciais.
