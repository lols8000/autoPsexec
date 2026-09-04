# Catálogo de módulos — Central N2 Workstation v5

## Convenções

- **Consulta:** somente leitura.
- **Remediação:** altera estado da estação.
- **Disruptiva:** pode interromper serviço, rede ou sessão.
- **Longa:** pode levar minutos.
- **Confirmação:** exigida quando a ação tem impacto relevante.

## Saúde / Compliance

**health.py** coleta hostname, usuário, Windows/build, fabricante/modelo, CPU, RAM, disco, uptime, reboot pendente, serviços automáticos parados, Defender, Firewall, GLPI, BitLocker, TPM e Secure Boot.

**compliance.py** compara o snapshot com o baseline ativo.

Tipo: consulta/análise.

## Performance

**performance.py** amostra CPU, RAM, disco e rede e identifica processos dominantes.

Tipo: consulta temporizada.

## Reparo do Windows

**repair.py**:

- SFC /scannow;
- DISM CheckHealth;
- DISM ScanHealth;
- DISM RestoreHealth;
- AnalyzeComponentStore;
- StartComponentCleanup;
- CHKDSK /scan;
- verificação WMI.

SFC/DISM/cleanup são operações pesadas. A v5 classifica essas ações para evitar concorrência incompatível por host.

## Dispositivos e drivers

**devices.py**:

- dispositivos PnP com erro;
- drivers assinados;
- USB presentes;
- pnputil /scan-devices;
- exportação de drivers.

### Inventário de drivers v5

O inventário:

- ignora registros completamente vazios;
- converte DriverDate para yyyy-MM-dd;
- agrupa entradas idênticas;
- inclui Count para representar instâncias repetidas;
- marca a visualização como drivers para a UI usar tabela específica.

A apresentação mostra Dispositivo, Fabricante, Versão, Data, Assinado, INF e Quantidade, além de resumo.

## Inicialização / Tarefas

**startup.py** consulta Win32_StartupCommand, chaves Run e serviços automáticos parados.

**tasks.py** lista tarefas, estado, última execução, próxima execução, resultado e falhas.

## Crashes / BSOD

**crashes.py** coleta BugCheck, Minidump, MEMORY.DMP e Application Error/Windows Error Reporting.

A Central localiza evidência; análise profunda de dump ainda pode exigir WinDbg.

## Segurança

**security.py** consulta Defender, proteção em tempo real, assinatura, Firewall, BitLocker, TPM, Secure Boot, RDP, SMBv1, UAC e ameaças recentes.

A Central não oferece botão genérico para desativar Defender/Firewall.

## Rede

**network.py**:

- adaptadores;
- IP/gateway/DNS;
- DHCP;
- flush DNS;
- renovação DHCP;
- reset Winsock/TCP-IP nas rotinas existentes;
- Wi-Fi;
- ARP;
- conexões TCP;
- teste TCP.

## Usuários / Perfis

**users_profiles.py** consulta administradores locais, perfis, SID, último uso e tamanho. Remoções, quando expostas, são destrutivas e exigem validação forte.

## Software

**software.py**:

- inventário por registro;
- disponibilidade do Winget;
- operações pelo catálogo configurado.

Winget pode ter comportamento diferente sob PsExec/SYSTEM.

## GLPI Agent

**glpi.py**:

- status;
- instalação/reparo;
- reinício de serviço;
- inventário forçado;
- log recente.

installer_source deve ficar em settings.local.json quando contiver infraestrutura interna.

## Impressoras

**printers.py**:

- inventário;
- filas;
- reinício do Spooler;
- limpeza de jobs.

Limpeza de fila remove documentos pendentes.

## Domínio / GPO

**domain.py**:

- domínio/membership;
- DC;
- Test-ComputerSecureChannel;
- w32tm;
- gpresult;
- reparo de secure channel;
- gpupdate.

## Disco e limpeza

**disk.py**:

- uso do C:;
- tamanho de perfis;
- estimativa de limpeza;
- limpeza segura de temporários e lixeira.

Downloads não são removidos automaticamente.

## Armazenamento / Bateria

**storage.py** usa Get-PhysicalDisk e classes WMI/CIM de bateria quando disponíveis.

Nem todo firmware/controlador expõe saúde detalhada.

## Ferramentas avançadas

**workstation_tools.py**:

- certificados;
- unidades mapeadas;
- shares locais;
- proxy;
- ativação/licenciamento;
- logons recentes.

## Sysinternals

**sysinternals.py** integra:

- Autorunsc;
- ProcDump;
- Handle;
- Sigcheck.

A Central não baixa a suíte automaticamente.

## Sistema

**system.py**:

- sessões;
- processos;
- serviços;
- start/stop/restart;
- GPUpdate;
- mensagem ao usuário;
- restart;
- shutdown;
- abort shutdown.

Ações de energia, kill de processo e parada de serviço são potencialmente disruptivas.

## Windows Update

**updates.py**:

- status/histórico;
- pendências;
- scan;
- reset de componentes.

## Pacote de diagnóstico

**diagnostic_package.py** agrega evidências para análise/escalonamento. Pacotes podem conter dados internos e não devem ser versionados.

## Jobs

**core/jobs.py** organiza execução, estado, concorrência e serialização de mutações por host.

## Diagnóstico / Correlação

**diagnostics/** normaliza Finding e Diagnosis e correlaciona sinais como pressão de armazenamento, degradação de storage, pressão de recursos, postura de segurança e falha do GLPI Agent.

## Playbooks

**playbooks/** define sequências de coleta orientadas por sintoma.

## Remediação

**remediation/** executa ações com snapshot before/after quando configurado.

## Persistência

**storage/database.py** persiste snapshots, jobs, findings, remediações e relatórios.

## Apresentação de resultados

A UI usa tabelas para listas de objetos quando possível. O JSON bruto fica reservado a estruturas complexas sem visualização específica.

A camada PsExec filtra ruído operacional em execuções bem-sucedidas sem esconder erro real.
