# Central N2 Workstation v5 — arquitetura consolidada

A v5 transforma o autoPsexec histórico em uma plataforma modular de troubleshooting para workstations Windows.

## Fluxo

~~~text
Alvo
 ↓
HostIdentity
 ↓
TransportManager
 ↓
Local / WinRM / PsExec
 ↓
Session + Connectivity + Capabilities
 ↓
Snapshot
 ↓
Findings
 ↓
Correlação
 ↓
Playbook
 ↓
Remediação
 ↓
Before / After
 ↓
Relatório / Histórico / GLPI
~~~

## Transportes

- **Local:** usado quando o alvo é a própria máquina.
- **WinRM:** preferido em destinos remotos quando utilizável.
- **PsExec:** fallback quando WinRM não está disponível e o executável existe.

WinRM não é obrigatório. Em redes onde 5985 está bloqueada, PsExec pode continuar operando por SMB/ADMIN$.

O executor descobre PsExec pelo PATH e pelos caminhos C:\Windows\System32\PsExec.exe e C:\Sysinternals\PsExec.exe.

## Saída humana

A interface v5 usa CommandResult.data para apresentar listas estruturadas em tabela quando possível.

O parser de JSON é tolerante a ruído do transporte. Em execuções PsExec bem-sucedidas, a camada de transporte remove mensagens operacionais sem valor para o suporte, como:

- início do processo remoto;
- encerramento com exit code 0;
- envelope CLIXML emitido pelo Windows PowerShell.

Erros reais são preservados.

O inventário de drivers agrupa entradas idênticas, normaliza datas, reduz registros vazios e apresenta assinatura em formato humano.

## Runtime e concorrência

JobManager mantém pool configurável. Leituras podem ser concorrentes; mutações por host são serializadas para evitar operações incompatíveis na mesma estação.

Estados de job:

- QUEUED
- RUNNING
- SUCCESS
- FAILED
- TIMEOUT
- CANCELLED

Timeout local não significa necessariamente cancelamento do processo remoto.

## Sessões

SessionManager mantém contexto lógico reutilizável por host: transporte, conectividade e capabilities. Não é um PSSession persistente do PowerShell.

## Capabilities

A Central detecta:

- versão do PowerShell;
- Windows/build;
- Winget;
- Defender;
- BitLocker;
- TPM;
- Secure Boot;
- Get-PhysicalDisk;
- bateria;
- GLPI.

Menus podem evitar consultas sem sentido, por exemplo bateria em desktop ou Winget ausente.

## Diagnóstico e playbooks

O motor separa fatos (Finding) de hipóteses (Diagnosis).

Playbooks disponíveis:

- computador lento;
- rede;
- impressão;
- domínio/GPO;
- Windows Update;
- crash de aplicação;
- BSOD;
- disco cheio;
- GLPI Agent.

## Remediação

RemediationEngine executa:

~~~text
snapshot antes
 ↓
ação
 ↓
snapshot depois
 ↓
diff
~~~

Remediações guiadas atuais:

- limpeza segura;
- reiniciar Spooler;
- reset de Windows Update;
- GPUpdate /force.

## Persistência

SQLite em data/central_n2.db mantém:

- hosts;
- snapshots;
- findings;
- remediações;
- jobs;
- relatórios.

O Diff compara snapshots recursivamente.

## Baselines

Perfis:

- DEFAULT;
- DESKTOP;
- NOTEBOOK;
- TI.

O compliance pode exigir disco livre, uptime, Defender, Firewall, GLPI, ausência de reboot pendente, BitLocker, TPM e Secure Boot.

## GLPI

settings.local.json é o local de configuração privada.

A API fica desabilitada por padrão. Quando habilitada, a Central pode enviar o relatório gerado como acompanhamento do chamado.

## Auditoria

Eventos podem carregar correlation_id. O logger mascara campos sensíveis como senha, token, Authorization, Bearer e API key.

## Interfaces

Console:

~~~powershell
python .\main.py
~~~

GUI:

~~~powershell
python .\main.py --gui
~~~

A GUI usa worker thread e não deve bloquear a janela durante operações remotas.

## Distribuição

- PyInstaller gera pacote portátil;
- Inno Setup gera instalador;
- workflow Windows executa compileall, pytest, build portátil e smoke do instalador;
- tags v* acionam workflow de release.

## Definition of Done

Novo recurso deve ter:

- tratamento de erro;
- timeout;
- retorno visual;
- resultado estruturado quando aplicável;
- logging;
- teste;
- documentação;
- comportamento compreensível para o técnico.
