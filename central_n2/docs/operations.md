# Manual operacional — Central N2 Workstation v5

## Princípio

~~~text
EVIDÊNCIA → DIAGNÓSTICO → REMEDIAÇÃO → VALIDAÇÃO → REGISTRO
~~~

Evite executar reparos apenas porque existem no menu.

## Inicialização

~~~powershell
python .\central_n2\main.py
~~~

A Central eleva via UAC quando necessário.

## Selecionar estação

Informe hostname ou IP. Hostname é preferível em ambientes de domínio.

Na seleção, a Central abre sessão lógica, identifica transporte/capabilities e coleta snapshot inicial.

## Interpretar transporte

### local

A própria máquina foi detectada. WinRM e PsExec são ignorados.

### winrm

PowerShell Remoting foi selecionado.

### psexec

WinRM não estava utilizável e PsExec foi selecionado.

## Cenário comum: WinRM bloqueado, PsExec disponível

Exemplo operacional real:

~~~text
Ping .............. OK
TCP 5985 .......... FAIL
TCP 445 ........... OK
ADMIN$ ............ OK
PsExec local ...... OK
~~~

Isso significa que a estação continua administrável por PsExec.

Teste:

~~~powershell
Test-NetConnection PC023 -Port 445
Test-Path \\PC023\ADMIN$
Test-Path C:\Sysinternals\PsExec.exe
C:\Sysinternals\PsExec.exe -accepteula \\PC023 hostname
~~~

## WinRM por IP

WinRM com autenticação padrão pode falhar quando o alvo é um IP. Em domínio, prefira hostname/FQDN.

Não amplie TrustedHosts para * indiscriminadamente.

## Saída da Central

Resultado:

~~~text
✓ SUCESSO [psexec] — 48060 ms
~~~

Listas estruturadas são mostradas em tabela quando possível.

A Central filtra ruído de PsExec/PowerShell em sucesso, mas preserva mensagens de erro.

## Drivers

Menu 5 → opção Drivers.

A saída mostra:

- dispositivo;
- fabricante;
- versão;
- data;
- assinatura;
- INF;
- quantidade agrupada.

Registros vazios são reduzidos e entradas idênticas são agrupadas.

## Saúde / Compliance

Use como triagem inicial. O score é heurística, não prova de causa.

## Performance

Use amostragem para lentidão. Observe CPU, RAM, disco, rede e processos dominantes.

## Reparo Windows

Sequência típica para suspeita de corrupção:

~~~text
CheckHealth → ScanHealth → RestoreHealth, se necessário → SFC
~~~

Não execute tudo por rotina.

## Hardware / Dispositivos

Use para Code 10/43, USB, driver suspeito e backup/exportação.

## Inicialização / Tarefas

Use em boot/login lento, persistência e tarefas corporativas com falha.

## Crashes / BSOD

Localize BugCheck, dumps e Application Error. Para análise profunda, escale com dump.

## Segurança

Use para postura, não para burlar controles.

## Rede

Diagnóstico recomendado:

~~~text
interface → IP → gateway → DNS → TCP específico
~~~

Flush DNS e DHCP têm impacto menor; reset de stack é mais agressivo.

## Usuários / Perfis

Antes de remoção de perfil, valide host, SID, sessão, dados e autorização.

## Software / GLPI Agent

Inventarie antes de instalar/remover. Winget pode diferir sob SYSTEM/PsExec.

## Impressoras

~~~text
inventário → fila → spooler → driver/rede
~~~

## Domínio / GPO

Verifique DC, hora, secure channel e gpresult antes de repair/gpupdate.

## Disco / Armazenamento

Use espaço, perfis, estimativa de limpeza e saúde física.

## Sysinternals

Ferramentas opcionais devem existir em diretório homologado.

## Pacote de diagnóstico

Use para escalonamento. Proteja o conteúdo.

## Energia / Processos / Serviços

Kill, stop, restart e shutdown podem causar perda de trabalho.

## Conectividade / Capabilities

Use quando a própria Central não consegue executar algo. Separe:

- DNS;
- ping;
- 5985/5986;
- ADMIN$;
- transporte selecionado;
- capabilities.

## Playbooks

Escolha o sintoma e deixe a Central coletar evidências correlacionadas.

## Histórico / Diff

Use para responder: “o que mudou desde o último atendimento?”.

## Relatório

Gere depois de diagnóstico/remediação. O correlation_id liga relatório, jobs e auditoria.

## Jobs

Consulte estado e tempo das operações.

## GLPI API

Gere relatório antes de enviar ao chamado.

## Remediações guiadas

Mostram impacto, exigem confirmação quando necessário e registram before/after.

## Baseline

Escolha DEFAULT, DESKTOP, NOTEBOOK ou TI conforme o tipo de estação.

## Timeout

Timeout local não garante encerramento remoto.

Antes de repetir operação pesada:

1. verifique se o processo continua;
2. consulte logs;
3. aguarde se necessário;
4. só então repita.

## Fluxos práticos

### Lentidão

~~~text
Saúde → Performance → Disco/Startup → Playbook lento → remediação se justificada
~~~

### Sem rede

~~~text
IP/Gateway/DNS → TCP → DHCP/flush apenas conforme causa
~~~

### Impressão

~~~text
Inventário → fila → Spooler → driver/rede
~~~

### Domínio

~~~text
DC → hora → secure channel → gpresult → repair/gpupdate se necessário
~~~

### WinRM indisponível

~~~text
5985 FAIL → 445 → ADMIN$ → PsExec → continuar atendimento
~~~
