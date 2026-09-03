# Catálogo de módulos — Central N2 Workstation v3

Este documento descreve as responsabilidades funcionais dos módulos da Central N2.

## Convenções

- **Consulta:** operação predominantemente somente leitura.
- **Remediação:** altera estado da estação.
- **Disruptiva:** pode interromper serviço, conectividade ou sessão.
- **Longa:** pode levar dezenas de segundos ou minutos.
- **Confirmação:** deve ser exigida pela UI quando a ação tiver impacto relevante.

---

## `health.py` — Saúde da estação

### Objetivo

Gerar snapshot rápido dos principais sinais de saúde de uma estação e alimentar o score visual.

### Dados coletados

- hostname;
- usuário;
- sistema operacional/build;
- fabricante/modelo;
- CPU;
- RAM;
- disco C:;
- uptime;
- reboot pendente;
- serviços automáticos parados;
- Defender;
- Firewall;
- GLPI Agent.

### Tipo

Consulta.

### Observação

O score é uma heurística operacional, não uma certificação de saúde do equipamento.

---

## `compliance.py` — Compliance

### Objetivo

Comparar o snapshot da estação com o baseline definido em `settings.json`.

### Regras atuais

- espaço livre mínimo;
- uptime máximo;
- Defender obrigatório;
- Firewall obrigatório;
- GLPI obrigatório;
- reboot pendente permitido ou não.

### Tipo

Consulta/análise local.

---

## `diagnostics.py` — Diagnóstico básico

### Objetivo

Executar preflight, inventário e verificações fundamentais antes de ações mais profundas.

### Uso recomendado

Primeiro passo quando o problema ainda não está classificado.

### Tipo

Consulta.

---

## `performance.py` — Performance

### Objetivo

Amostrar a estação por um período e evitar diagnósticos baseados em um único valor instantâneo.

### Coleta

- CPU total;
- RAM usada;
- disco ativo;
- bytes de rede por segundo;
- top processos por CPU;
- top processos por memória.

### Tipo

Consulta temporizada.

### Cuidado

Amostragens maiores aumentam tempo de execução. Não é um monitoramento contínuo.

---

## `repair.py` — Reparo do Windows

### Operações

- `sfc /scannow`;
- DISM `/CheckHealth`;
- DISM `/ScanHealth`;
- DISM `/RestoreHealth`;
- DISM `/AnalyzeComponentStore`;
- DISM `/StartComponentCleanup`;
- `chkdsk C: /scan`;
- `winmgmt /verifyrepository`.

### Tipo

Consulta e remediação.

### Risco

SFC/DISM/Component Cleanup podem ser longos e consumir CPU/disco. Devem ser executados conscientemente, preferencialmente sem outras manutenções pesadas simultâneas na mesma estação.

---

## `devices.py` — Dispositivos e drivers

### Operações

- dispositivos PnP com erro;
- inventário de drivers assinados;
- dispositivos USB presentes;
- `pnputil /scan-devices`;
- exportação de drivers com `pnputil /export-driver`.

### Tipo

Consulta e remediação leve.

### Observação

Exportação de drivers usa caminho local da estação alvo.

---

## `startup.py` — Inicialização

### Coleta

- `Win32_StartupCommand`;
- chaves `Run` de HKLM/HKCU;
- serviços automáticos parados.

### Tipo

Consulta.

### Uso

Diagnóstico de lentidão, software persistente e itens de inicialização inesperados.

---

## `tasks.py` — Tarefas agendadas

### Operações

- inventário de tarefas;
- estado;
- última execução;
- próximo agendamento;
- resultado da última execução;
- filtro de tarefas com falha.

### Tipo

Consulta.

---

## `crashes.py` — Crashes e BSOD

### Coleta

- BugCheck em System Event Log;
- arquivos de `C:\Windows\Minidump`;
- `C:\Windows\MEMORY.DMP`;
- eventos Application Error/Windows Error Reporting.

### Tipo

Consulta.

### Limitação

O módulo identifica evidências; ele não substitui WinDbg para análise profunda de dump.

---

## `security.py` — Segurança da estação

### Postura consultada

- Microsoft Defender;
- proteção em tempo real;
- atualização de assinatura;
- Firewall;
- BitLocker;
- TPM;
- Secure Boot;
- RDP;
- SMBv1;
- UAC;
- ameaças recentes.

### Tipo

Consulta.

### Política

Não deve oferecer ações de desativação de Defender ou Firewall.

---

## `network.py` — Rede da estação

### Operações

- adaptadores;
- configuração IP;
- gateway;
- DNS;
- renovação DHCP;
- flush DNS;
- reset Winsock;
- reset TCP/IP;
- Wi-Fi;
- ARP;
- conexões TCP;
- teste TCP.

### Tipo

Consulta e remediação.

### Risco

Reset TCP/IP/Winsock e alteração de Wi-Fi podem afetar conectividade e exigir reboot.

---

## `users_profiles.py` — Usuários e perfis

### Operações

- administradores locais;
- perfis;
- tamanho de perfil;
- SID;
- último uso;
- remoção de perfil quando solicitada.

### Tipo

Consulta e remediação destrutiva.

### Regra

Remoção de perfil exige confirmação forte e validação de SID/host correto.

---

## `software.py` — Software

### Operações

- inventário pelo registro;
- presença do Winget;
- instalação pelo catálogo;
- upgrade pelo catálogo;
- remoção pelo catálogo.

### Tipo

Consulta e remediação.

### Limitação

Winget pode não funcionar no mesmo contexto quando a execução remota ocorre como SYSTEM/PsExec. Esse comportamento depende do ambiente.

---

## `glpi.py` — GLPI Agent

### Operações

- status;
- instalação/reparo;
- reinício do serviço;
- inventário forçado;
- consulta de log recente.

### Dependência de configuração

`glpi.installer_source` deve ser configurado para o ambiente.

### Segurança

O repositório público mantém esse valor vazio por padrão.

---

## `printers.py` — Impressoras

### Operações

- inventário de impressoras;
- filas;
- reinício do Spooler;
- limpeza de jobs.

### Tipo

Consulta e remediação.

### Risco

Limpar fila remove jobs pendentes. Deve ser uma decisão consciente do suporte.

---

## `domain.py` — Domínio e GPO

### Operações

- domínio e membership;
- Domain Controller;
- `Test-ComputerSecureChannel`;
- `w32tm /query /status`;
- `gpresult`;
- reparo do secure channel;
- `gpupdate` é reutilizado pelo módulo de sistema/UI.

### Tipo

Consulta e remediação.

### Risco

Reparo de secure channel exige contexto administrativo adequado no domínio.

---

## `disk.py` — Disco e limpeza

### Operações

- uso do C:;
- tamanho de perfis;
- estimativa de espaço recuperável;
- limpeza de TEMP e lixeira;
- limpeza segura sem remover Downloads automaticamente.

### Tipo

Consulta e remediação.

### Política

A ferramenta não deve apagar pastas de usuário arbitrariamente como parte de uma limpeza genérica.

---

## `storage.py` — Armazenamento físico e bateria

### Operações

- `Get-PhysicalDisk`;
- serial;
- tipo de mídia;
- bus;
- HealthStatus;
- OperationalStatus;
- tamanho;
- informações de bateria via classes WMI/CIM quando disponíveis.

### Tipo

Consulta.

### Limitação

Nem todo firmware/controlador expõe dados completos de saúde ou bateria.

---

## `workstation_tools.py` — Ferramentas avançadas

### Operações

- certificados `LocalMachine\My`;
- unidades de rede;
- compartilhamentos SMB locais;
- proxy WinHTTP e Internet Settings;
- ativação/licenciamento Windows;
- logons recentes no Event Log de Security.

### Tipo

Consulta.

---

## `sysinternals.py` — Integração Sysinternals

### Dependência

Diretório configurado em:

```json
"sysinternals_dir": "C:\\Sysinternals"
```

### Ferramentas utilizadas

#### Autorunsc

Inventário avançado de persistência/startup.

#### ProcDump

Captura de dump de processo.

#### Handle

Localização de processo que mantém handle aberto para arquivo/recurso.

#### Sigcheck

Assinatura, hash e informações do executável.

### Tipo

Consulta e coleta; ProcDump gera arquivo na estação.

### Política

A Central não baixa executáveis Sysinternals automaticamente.

---

## `system.py` — Sistema

### Operações

- sessões;
- processos;
- finalizar processo;
- serviços;
- start/stop/restart de serviço;
- GPUpdate;
- mensagem ao usuário;
- restart;
- shutdown;
- abort shutdown.

### Tipo

Consulta e remediação/disruptiva.

### Risco

Finalização de processo, stop de serviço e energia exigem atenção e confirmação apropriada.

---

## `updates.py` — Windows Update

### Operações

- status/histórico;
- pendências;
- trigger de scan;
- reset de componentes do Windows Update.

### Tipo

Consulta e remediação.

---

## `diagnostic_package.py` — Pacote de diagnóstico

### Objetivo

Concentrar diferentes evidências em um artefato local para análise, escalonamento ou documentação de chamado.

### Tipo

Coleta.

### Segurança

Pacotes podem conter hostname, usuário, software, endereços e outras informações do ambiente. A pasta de relatórios não deve ser versionada.

---

## `batch.py` — Lote

### Objetivo

Executar operações em múltiplos hosts com limite de concorrência.

### Regra

Ações em lote devem permanecer restritas a operações compatíveis com execução concorrente e com impacto conhecido.

Nunca transformar o módulo de lote em um mecanismo de execução destrutiva massiva sem controles adicionais.
