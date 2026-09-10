# Catálogo de módulos — Central N2 Workstation 5.1.0

## Convenções

- **READ_ONLY**: consulta leve;
- **HEAVY_READ**: consulta custosa;
- **LIGHT_WRITE**: alteração de baixo impacto;
- **HEAVY_WRITE**: alteração pesada;
- **DISRUPTIVE**: pode afetar sessão/rede/energia.

Operações não READ_ONLY são serializadas por host pelo JobManager.

## Saúde / Compliance

\`health.py\` coleta host, usuário, Windows/build, hardware, CPU, RAM, disco, uptime, reboot pendente, serviços automáticos, Defender, Firewall, GLPI, BitLocker, TPM e Secure Boot.

\`compliance.py\` usa o motor comum de avaliação.

Estados: PASS, FAIL, UNKNOWN e NOT_APPLICABLE.

## Performance

\`performance.py\` amostra CPU, RAM, disco e rede e identifica processos dominantes.

## Reparo Windows

\`repair.py\` expõe SFC, DISM, Component Store, CHKDSK e WMI.

Operações pesadas usam classes de job compatíveis com serialização por host.

## Dispositivos / Drivers

\`devices.py\`:

- PnP com erro;
- inventário de drivers;
- USB;
- rescan;
- exportação.

Inventário de drivers normaliza datas, agrupa registros equivalentes e mostra assinatura/INF/contagem.

## Inicialização / Tarefas

\`startup.py\`: startup commands, chaves Run e serviços automáticos parados.

\`tasks.py\`: tarefas agendadas, estado, última/próxima execução e falhas.

## Crashes / BSOD

\`crashes.py\`: BugCheck, Minidump, MEMORY.DMP e Application Error/WER.

## Segurança

\`security.py\`: Defender, Firewall, BitLocker, TPM, Secure Boot, RDP, SMBv1, UAC e ameaças.

A Central não fornece ação genérica para desligar controles de segurança.

## Rede

\`network.py\`:

- adaptadores;
- IP/gateway/DNS;
- DHCP;
- ARP;
- conexões;
- flush DNS;
- renovação DHCP;
- resets existentes.

Mutações de rede usam o caminho seguro de execução mutável.

## Usuários / Perfis

\`users_profiles.py\` consulta administradores locais, perfis, SID, último uso e tamanho.

Ações destrutivas devem exigir validação/confirmar alvo e nunca ser generalizadas em lote sem controle.

## Software / Winget

\`software.py\`:

- inventário por registro;
- disponibilidade do Winget;
- install/upgrade/uninstall pelo catálogo permitido.

Operações Winget verificam \`$LASTEXITCODE\`.

## GLPI Agent

\`glpi.py\`:

- status;
- cópia de instalador homologado;
- instalação/reparo;
- reinício de serviço;
- inventário forçado;
- log recente.

## GLPI API

\`integrations/glpi/client.py\` encapsula sessão, erros HTTP/rede, JSON e follow-up de ticket.

## Impressoras

\`printers.py\`:

- inventário;
- fila;
- status do Spooler;
- reinício;
- limpeza de fila.

A remediação guiada de Spooler valida estado Running depois da ação.

## Domínio / GPO

\`domain.py\`:

- status de domínio;
- DC;
- secure channel;
- horário;
- gpresult;
- gpupdate;
- repair de secure channel.

## Disco

\`disk.py\`:

- uso do C:;
- perfis por tamanho;
- estimativa de limpeza;
- limpeza segura.

A limpeza segura atua em \`%TEMP%\` e \`%SystemRoot%\Temp\`; **não toca Lixeira, Downloads ou cache do Windows Update**.

## Armazenamento / Bateria

\`storage.py\` consulta Get-PhysicalDisk e WMI/CIM de bateria quando disponíveis.

## Ferramentas avançadas

\`workstation_tools.py\` consulta certificados, unidades mapeadas, shares, proxy, ativação e logons.

## Sysinternals

\`sysinternals.py\`: Autorunsc, ProcDump, Handle e Sigcheck.

Não há download automático.

## Sistema

\`system.py\`:

- sessões;
- processos;
- serviços;
- GPUpdate;
- mensagens;
- restart/shutdown/abort.

Ações de energia e mudança de serviço são classificadas como mutações/disruptivas.

## Windows Update

\`updates.py\`:

- status/histórico;
- pendências;
- scan;
- reset transacional de componentes.

No reset, \`SoftwareDistribution\` e \`catroot2\` são renomeados com timestamp e serviços originalmente ativos são restaurados em bloco \`finally\`.

## Pacote diagnóstico

\`diagnostic_package.py\` agrega evidências para escalonamento.

Pacotes são dados operacionais; não versionar.

## Conectividade

\`core/connectivity.py\` avalia DNS, ping, 445, 5985, 5986, WinRM autenticado, ADMIN$ e PsExec real.

## Executor

\`core/executor.py\` concentra semântica de transporte e fallback seguro.

Use APIs mutáveis para qualquer ação que altere o host.

## Jobs / Batch

\`core/jobs.py\` é o scheduler central.

\`modules/batch.py\` pode reutilizar esse mesmo scheduler, evitando pools paralelos independentes.

## Diagnóstico / Correlação

\`diagnostics/\` separa Finding de Diagnosis.

## Playbooks

\`playbooks/\` coleta evidências orientadas por sintoma. Não aplica remediação silenciosa.

## Remediação

\`remediation/\` executa before/action/after/validator.

Validadores atuais: limpeza, Spooler, Windows Update e GPUpdate.

## Persistência

\`storage/database.py\` usa migrations versionadas e persiste correlation_id.

## Relatórios

\`reports/\` gera Markdown, JSON e TXT; nomes de arquivos são sanitizados.
