# Central N2 Workstation — Guia rápido v5

**Versão:** 5.0.0  
**Console:** python main.py  
**GUI opcional:** python main.py --gui

## Visão rápida

A v5 detecta automaticamente quando o alvo é a própria estação e executa localmente. Para destinos remotos, WinRM é preferido e PsExec é fallback.

~~~text
LOCAL → execução direta

REMOTO
  ↓
WinRM utilizável?
  ├─ SIM → WinRM
  └─ NÃO → PsExec, se disponível
~~~

WinRM não é requisito absoluto. Um ambiente com 5985 bloqueada ainda pode ser administrado por PsExec se SMB/445 e ADMIN$ estiverem disponíveis.

## Requisitos

- Windows 10/11 ou Windows Server como estação administrativa;
- Python 3.10+ para execução por código-fonte;
- privilégio administrativo;
- conectividade com o host alvo;
- WinRM configurado, quando usado;
- PsExec disponível quando for necessário fallback;
- Sysinternals opcional.

## PsExec recomendado

Diretório operacional recomendado:

~~~text
C:\Sysinternals\
~~~

Validação:

~~~powershell
Test-Path C:\Sysinternals\PsExec.exe
Get-Command PsExec.exe -ErrorAction SilentlyContinue
~~~

Teste remoto:

~~~powershell
C:\Sysinternals\PsExec.exe -accepteula \\PC023 hostname
~~~

A Central também procura PsExec em C:\Windows\System32 e no PATH.

## Configuração local

Copie:

~~~text
config\settings.local.example.json
~~~

para:

~~~text
config\settings.local.json
~~~

O ConfigLoader aplica merge recursivo sobre settings.json. Nunca versione segredos.

## Iniciar

Da raiz do repositório:

~~~powershell
python .\central_n2\main.py
~~~

Ou:

~~~powershell
cd central_n2
python .\main.py
~~~

A aplicação solicita UAC quando necessário.

## Menu v5

~~~text
[1]  Selecionar estação
[2]  Saúde / Compliance
[3]  Performance
[4]  Reparo do Windows
[5]  Hardware / Drivers / Dispositivos
[6]  Inicialização / Tarefas
[7]  Crashes / BSOD
[8]  Segurança
[9]  Rede
[10] Usuários / Perfis
[11] Software / GLPI Agent
[12] Impressoras
[13] Domínio / GPO
[14] Disco / Armazenamento / Bateria
[15] Ferramentas avançadas
[16] Sysinternals
[17] Pacote de diagnóstico
[18] Energia / Processos / Serviços
[19] Conectividade / Capabilities
[20] Assistente N2 / Playbooks
[21] Histórico / Diff
[22] Gerar relatório
[23] Jobs
[24] Atualização da Central
[25] GLPI API
[26] Remediações guiadas
[27] Perfil / Baseline
[0]  Sair
~~~

## Retorno visual

Exemplo:

~~~text
✓ SUCESSO [psexec] — 48060 ms
~~~

Resultados estruturados são formatados como tabela quando possível.

### Exemplo: drivers

~~~text
Dispositivo                         | Fabricante              | Versão           | Data       | Assinado | INF
------------------------------------+-------------------------+------------------+------------+----------+----------
AMD Radeon(TM) Graphics             | Advanced Micro Devices  | 31.0.12027.9001  | 2023-03-30 | SIM      | oem53.inf
AMD High Definition Audio Device    | Advanced Micro Devices  | 10.0.1.38        | 2024-04-26 | SIM      | oem48.inf

Resumo: 2 instâncias | 2 entradas agrupadas | 0 não assinada(s)
~~~

No PsExec, mensagens técnicas como Starting powershell.exe, CLIXML e exit code 0 são removidas quando são apenas ruído de uma execução bem-sucedida. Falhas reais não são escondidas.

## Diagnóstico de transporte

Quando WinRM falhar:

~~~powershell
Resolve-DnsName PC023
Test-NetConnection PC023 -Port 5985
Test-NetConnection PC023 -Port 445
Test-Path \\PC023\ADMIN$
Test-Path C:\Sysinternals\PsExec.exe
~~~

Interpretação típica:

~~~text
Ping/SMB/ADMIN$ OK + WinRM 5985 FAIL + PsExec disponível
→ estação administrável por PsExec
~~~

## Operação segura

Fluxo recomendado:

~~~text
Saúde
 ↓
Diagnóstico / Playbook
 ↓
Confirmar causa
 ↓
Remediar
 ↓
Validar antes/depois
 ↓
Gerar relatório / histórico
~~~

Timeout local não garante que um processo remoto já iniciado tenha sido encerrado.
