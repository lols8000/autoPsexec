# Central de Manutenção N2 — Estações Windows

Central modular para diagnóstico, manutenção e troubleshooting remoto de estações Windows. A v3 mantém o foco exclusivamente em workstation e adiciona uma camada responsiva para que operações longas não deixem o suporte sem feedback visual.

## Destaques da v3

- execução de operações bloqueantes em worker threads;
- heartbeat visual contínuo com tempo decorrido;
- timeout padronizado para evitar espera indefinida;
- resultado final sempre exibido como sucesso, falha ou timeout;
- encerramento controlado do pool de workers;
- módulos novos de reparo do Windows, performance, drivers/dispositivos, startup, tarefas agendadas, crashes/BSOD, armazenamento/bateria e ferramentas avançadas;
- integração opcional com Sysinternals já instalado na estação;
- configuração pública sanitizada, sem caminho interno do ambiente versionado.

## Execução

```powershell
python .\central_n2\main.py
```

A aplicação solicita elevação UAC quando necessário.

## Feedback visual

Toda ação iniciada pela interface v3 passa pelo `ResponsiveJobRunner`.

Durante a execução o suporte vê algo semelhante a:

```text
| DISM RestoreHealth...   4.2s
/ DISM RestoreHealth...   8.7s
- DISM RestoreHealth...  15.3s
```

Ao concluir:

```text
✓ SUCESSO [winrm] — 18432 ms
```

Em falha:

```text
✗ FALHA [psexec]
Erro: ...
```

Se exceder o limite configurado:

```text
✗ TIMEOUT: Operação '...' excedeu ...s
```

O heartbeat não significa que o comando remoto está produzindo saída a cada instante; ele confirma que a Central continua viva enquanto aguarda o worker remoto.

## Concorrência

A v3 usa `ThreadPoolExecutor` para operações I/O-bound e remotas. Isso faz sentido porque a maior parte do tempo é gasta aguardando WinRM, PsExec, PowerShell, rede ou processos do Windows, e não consumindo CPU Python.

A interface usa threads para manter responsividade. Ações em lote continuam com concorrência limitada para não sobrecarregar rede ou endpoints.

## Menu v3

```text
[1] Selecionar estação
[2] Saúde / Compliance
[3] Performance
[4] Reparo do Windows
[5] Hardware / Drivers / Dispositivos
[6] Inicialização / Tarefas
[7] Crashes / BSOD
[8] Segurança
[9] Rede
[10] Usuários / Perfis
[11] Software / GLPI
[12] Impressoras
[13] Domínio / GPO
[14] Disco / Armazenamento / Bateria
[15] Ferramentas avançadas
[16] Sysinternals
[17] Pacote de diagnóstico
[18] Energia / Processos / Serviços
```

## Reparo do Windows

Disponível com retorno visual e timeout longo:

- `sfc /scannow`;
- DISM CheckHealth;
- DISM ScanHealth;
- DISM RestoreHealth;
- análise e limpeza do Component Store;
- CHKDSK online;
- verificação de consistência do repositório WMI.

Ações de maior impacto continuam exigindo confirmação.

## Performance

A Central pode coletar amostras por alguns segundos e retornar:

- CPU média e pico;
- RAM utilizada;
- atividade média de disco;
- throughput de rede aproximado;
- processos com maior CPU;
- processos com maior uso de memória.

A coleta usa amostragem limitada e não mantém monitoramento permanente em background.

## Drivers e dispositivos

- dispositivos com erro no Plug and Play;
- código de erro do Device Manager;
- inventário de drivers assinados;
- dispositivos USB presentes;
- `pnputil /scan-devices`;
- exportação de drivers.

## Inicialização e tarefas

- Run/Startup via WMI e Registro;
- tarefas agendadas;
- tarefas com último resultado diferente de zero;
- serviços automáticos parados.

## Crashes e BSOD

- eventos BugCheck;
- conteúdo de `C:\Windows\Minidump`;
- presença de `MEMORY.DMP`;
- eventos de crash de aplicações.

## Segurança

- Defender;
- proteção em tempo real;
- assinatura do antivírus;
- Firewall;
- BitLocker;
- TPM;
- Secure Boot;
- RDP;
- SMBv1;
- UAC;
- ameaças recentes.

A Central não oferece botão para desativar Defender ou Firewall.

## Armazenamento e bateria

- volumes e espaço livre;
- discos físicos;
- status/saúde quando exposto pelo Windows;
- tipo de mídia e barramento;
- bateria;
- capacidade projetada e carga total quando suportado;
- cálculo de saúde da bateria;
- ciclos quando o firmware expõe a informação.

## Ferramentas avançadas

- certificados da máquina e certificados expirando;
- unidades mapeadas;
- compartilhamentos locais;
- proxy WinHTTP/usuário;
- status de ativação do Windows;
- logons recentes.

## Sysinternals opcional

Por padrão a Central procura em:

```text
C:\Sysinternals
```

O diretório é configurável em `config/settings.json`.

Ferramentas integradas quando presentes:

- `autorunsc.exe` — inventário avançado de inicialização;
- `procdump.exe` — captura de dump de processo;
- `handle.exe` — localizar arquivo/objeto aberto;
- `sigcheck.exe` — assinatura e hash;
- inventário de PsPing, PsLoggedOn, RAMMap e Procmon para expansão futura.

A Central não baixa executáveis automaticamente.

## Configuração

Arquivo principal:

```text
central_n2/config/settings.json
```

Parâmetros relevantes da v3:

```json
{
  "sysinternals_dir": "C:\\Sysinternals",
  "ui": {
    "heartbeat_seconds": 0.2,
    "long_operation_timeout_seconds": 3600
  }
}
```

O `glpi.installer_source` versionado fica vazio de propósito. Configure o caminho apropriado ao ambiente antes de usar instalação/reparo do agente.

## Testes

```powershell
cd central_n2
python -m pip install pytest
python -m pytest -q
```

O workflow do GitHub Actions também executa:

- compilação de todos os arquivos Python;
- testes centrais;
- saúde/compliance;
- runner responsivo;
- heartbeat;
- timeout;
- compatibilidade da API dos módulos usados pela interface v3.

## Princípios de operação

1. Toda operação longa precisa mostrar que a Central continua viva.
2. A interface não deve congelar aguardando I/O remoto.
3. Timeout deve existir em toda operação potencialmente bloqueante.
4. Ações destrutivas continuam exigindo confirmação explícita.
5. Threads são usadas para I/O e operações remotas; não para paralelizar indiscriminadamente tarefas destrutivas.
6. A ferramenta deve retornar evidência suficiente para o suporte registrar o atendimento.
7. Falha de ferramenta opcional não deve derrubar a Central inteira.
