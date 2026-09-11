# Catálogo de módulos — Central N2 Workstation 5.2.0

## Convenções

OperationClass:

- READ_ONLY;
- HEAVY_READ;
- LIGHT_WRITE;
- HEAVY_WRITE;
- DISRUPTIVE.

Mutações são serializadas por host.

## Módulos de diagnóstico

### health.py

Snapshot de host, Windows, hardware, CPU, RAM, disco, uptime, reboot pending, serviços, Defender, Firewall, GLPI, BitLocker, TPM e Secure Boot.

### performance.py

CPU/RAM/disco/rede e processos dominantes.

### crashes.py

BugCheck, minidumps, MEMORY.DMP, Application Error/WER.

### storage.py

PhysicalDisk e bateria.

### startup.py / tasks.py

Startup, chaves Run, serviços automáticos e Scheduled Tasks.

## Módulos de execução

### system.py

- sessões;
- processos;
- status/ações de serviço;
- StartType;
- GPUpdate;
- mensagem;
- logoff;
- restart/shutdown/abort.

### network.py

- adaptadores;
- IP/gateway/DNS;
- DHCP;
- flush/register DNS;
- Winsock/TCP-IP;
- ARP;
- enable/disable/restart de adaptador.

Restart de NIC usa workflow de disconnect temporário na Central de Execuções.

### printers.py

- inventário;
- fila;
- Spooler;
- limpar fila;
- adicionar conexão;
- remover impressora.

### devices.py

- PnP com erro;
- PnP presente;
- drivers;
- USB;
- enable/disable device;
- rescan;
- install/remove driver package;
- export.

### users_profiles.py

- admins locais;
- profiles;
- profile status;
- limpeza de TEMP por SID;
- remoção controlada de perfil.

### software.py

- inventário;
- Winget;
- install/upgrade/uninstall pelo catálogo.

### glpi.py

- status;
- instalação/reparo;
- restart;
- force inventory;
- logs.

### security.py

- posture;
- threats;
- Defender status;
- signature update;
- quick/full scan.

### domain.py

- domínio/DC;
- secure channel;
- gpresult/gpupdate;
- w32time;
- resync;
- purge Kerberos SYSTEM.

### updates.py

- histórico/status;
- scan;
- install pending;
- reset transacional dos componentes.

### repair.py

- SFC;
- DISM;
- Component Store;
- CHKDSK;
- WMI;
- Store reset.

### disk.py

- uso;
- profiles por tamanho;
- estimativa;
- cleanup seguro.

### packages.py

Instala apenas pacote homologado em configuração.

### certificates.py

Importa apenas certificado público homologado.

### registry_actions.py

Executa e inspeciona apenas RegistryAction homologada; suporta rollback do valor anterior.

### file_ops.py

- ensure directory;
- path status;
- move/rename;
- remove file.

Não oferece delete recursivo genérico.

## Catálogos de execução

`execution/catalogs/` separa definição operacional do módulo técnico.

Cada catálogo liga:

```text
ExecutionAction
+ handler
+ before_probe
+ after_probe
+ validator
+ preconditions
+ rollback opcional
```

## Core de execução

### execution/models.py

Contratos de ação, parâmetro, risco, retry, disconnect, selector, plan, record e recovery.

### execution/registry.py

Registro, invariantes, categorias e busca.

### execution/policy.py

Transport, privilege, capabilities e preconditions.

### execution/engine.py

Plan, execute, recovery, validation e rollback.

### execution/config_validation.py

Sanitização das allowlists no bootstrap.

## Persistência

`storage/database.py` usa schema 4 e persiste executions/rollbacks com correlation_id.

## Relatórios

`reports/` gera Markdown/JSON/TXT.

## Regra de módulo

Módulo não conhece UI, não chama `input()` e retorna `CommandResult`.
