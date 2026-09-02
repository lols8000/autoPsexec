# Central Remota de Manutenção N2

Evolução do projeto histórico `autoPsexec` para uma central modular de administração remota de estações Windows. Os scripts antigos da raiz do repositório são preservados como referência e não são utilizados pela nova aplicação.

## O que já está implementado

- seleção persistente do computador alvo;
- pré-flight automático com DNS, ping, `ADMIN$`, WinRM e inventário básico;
- PowerShell Remoting/WinRM como transporte preferencial;
- PsExec/Sysinternals como fallback;
- execução sem `shell=True` no motor principal;
- inventário por PowerShell/CIM, sem dependência do WMIC;
- informações de fabricante, modelo, serial, BIOS, Windows, CPU, RAM, disco e uptime;
- consulta de eventos críticos e reinicialização pendente;
- sessões de usuário (`quser`);
- processos e finalização controlada;
- serviços: listar, iniciar, parar e reiniciar;
- `gpupdate /force` e envio de mensagens;
- rede: interfaces, MAC, IP, gateway, DNS, DHCP, cache DNS, Winsock, TCP/IP, Wi-Fi, ARP e conexões TCP;
- teste TCP partindo do computador remoto;
- catálogo de software com Winget;
- listar, instalar, atualizar e remover software homologado;
- GLPI Agent: status, instalação/reparo, serviço, inventário forçado e leitura de log;
- reiniciar/desligar com confirmação e cancelamento de shutdown agendado;
- execução em lote concorrente;
- auditoria em JSONL por operador/host/ação;
- testes unitários básicos.

## Estrutura

```text
central_n2/
├── main.py
├── pyproject.toml
├── config/
│   └── settings.json
├── core/
│   ├── executor.py
│   ├── logger.py
│   └── result.py
├── modules/
│   ├── batch.py
│   ├── diagnostics.py
│   ├── glpi.py
│   ├── network.py
│   ├── software.py
│   └── system.py
├── ui/
│   └── console.py
├── tests/
│   └── test_core.py
└── logs/
    └── YYYY-MM-DD.jsonl   # criado em execução
```

## Requisitos

- Windows 10/11 ou Windows Server na estação administrativa;
- Python 3.10 ou superior;
- privilégios administrativos;
- conectividade e permissões administrativas para os computadores alvo;
- PowerShell Remoting/WinRM configurado para o caminho preferencial;
- PsExec disponível para fallback, quando necessário.

A aplicação usa somente a biblioteca padrão do Python em produção. `pytest` é dependência opcional de desenvolvimento.

## Execução

Abra o PowerShell ou Prompt na raiz do repositório e execute:

```powershell
python .\central_n2\main.py
```

Se o processo não estiver elevado, a aplicação solicitará UAC e reiniciará como administrador.

## Configuração

Edite:

```text
central_n2/config/settings.json
```

Exemplo dos principais parâmetros:

```json
{
  "timeout_seconds": 60,
  "psexec_path": "C:\\Windows\\System32\\PsExec.exe",
  "glpi": {
    "installer_source": "\\\\servidor\\share\\glpiagentinstall.vbs",
    "remote_installer_path": "C:\\glpiagentinstall.vbs"
  }
}
```

Se o caminho configurado do PsExec não existir, o executor tenta encontrá-lo no `PATH`, em `C:\Windows\System32` e em `C:\Sysinternals`.

### Catálogo de software

O catálogo também fica no `settings.json`:

```json
"software": {
  "chrome": {"name": "Google Chrome", "winget_id": "Google.Chrome"},
  "firefox": {"name": "Mozilla Firefox", "winget_id": "Mozilla.Firefox"}
}
```

Isso evita deixar IDs de pacote espalhados pelo código e facilita homologação centralizada.

## WinRM e fallback

O fluxo de execução é:

```text
Ação solicitada
      │
      ▼
Test-WSMan
  │        │
  OK      falha
  │        │
WinRM    PsExec
  │        │
  └────┬───┘
       ▼
CommandResult
       │
       ├── saída
       ├── erro
       ├── código
       ├── duração
       └── transporte
```

Não é recomendado habilitar WinRM indiscriminadamente em redes não confiáveis. Em domínio, prefira configurar PowerShell Remoting, firewall e autenticação por política corporativa/GPO.

## Auditoria

Cada execução relevante gera um registro em:

```text
central_n2/logs/YYYY-MM-DD.jsonl
```

Exemplo:

```json
{
  "timestamp": "2026-09-02T18:31:04-04:00",
  "operator": "MBezerra",
  "action": "powershell_remote",
  "host": "PC-ADM-023",
  "success": true,
  "transport": "winrm",
  "duration_ms": 1542
}
```

Não armazene senhas, tokens ou credenciais no `settings.json` ou nos logs.

## Operações sensíveis

A interface exige confirmação textual para operações com maior impacto, como:

- finalizar processos;
- parar/reiniciar serviços;
- resetar Winsock/TCP-IP;
- desabilitar Wi-Fi;
- remover software;
- reiniciar ou desligar computadores.

Ainda assim, a ferramenta deve ser usada somente por administradores autorizados e primeiro testada em um grupo pequeno de máquinas.

## Ações em lote

O menu aceita:

```text
PC001,PC002,PC003
```

ou um arquivo `.txt` com um host por linha.

A primeira versão permite em lote:

- ping;
- GPUpdate;
- flush DNS;
- inventário GLPI.

A concorrência padrão é limitada a 5 hosts para evitar gerar carga desnecessária na infraestrutura.

## Testes

Instale a dependência opcional:

```powershell
python -m pip install pytest
```

Depois:

```powershell
cd central_n2
python -m pytest
```

Os testes atuais cobrem resultado padronizado, auditoria e executor de lote sem exigir uma máquina remota.

## Limitações conhecidas da v1

- operações Winget dependem de Winget funcional no contexto remoto;
- ambiente sem WinRM depende de `ADMIN$`, serviço RPC/SMB e PsExec permitido;
- o instalador GLPI atualmente segue o VBS utilizado no ambiente original e deve ser ajustado no `settings.json` antes de uso em outro local;
- o módulo de switch/MAC/OUI discutido para Intelbras não foi misturado nesta primeira versão: ele deve entrar como módulo separado para não acoplar administração de endpoints à configuração de rede;
- a aplicação ainda é console/TUI; uma GUI pode ser construída sobre os mesmos módulos sem reescrever o motor.

## Princípios do projeto

1. preservar o código histórico;
2. separar interface, regra e transporte;
3. não depender de offsets frágeis de saída textual para inventário;
4. preferir PowerShell/CIM moderno ao WMIC;
5. manter PsExec como fallback, não como arquitetura inteira;
6. registrar ações administrativas;
7. confirmar operações disruptivas;
8. permitir expansão modular sem transformar `main.py` em um monólito.
