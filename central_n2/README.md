# Central Remota de Manutenção N2

Evolução do projeto histórico `autoPsexec` para uma central modular de administração remota de estações Windows e auditoria de acesso físico à rede. Os scripts antigos da raiz do repositório são preservados como referência e não são utilizados pela nova aplicação.

## O que já está implementado

### Endpoints Windows

- seleção persistente do computador alvo;
- pré-flight automático com DNS, ping, `ADMIN$`, WinRM e inventário básico;
- PowerShell Remoting/WinRM como transporte preferencial;
- PsExec/Sysinternals como fallback;
- execução sem `shell=True` no motor principal;
- inventário por PowerShell/CIM, sem dependência do WMIC;
- fabricante, modelo, serial, BIOS, Windows, CPU, RAM, disco e uptime;
- eventos críticos e reinicialização pendente;
- sessões de usuário, processos e serviços;
- `gpupdate /force` e mensagens remotas;
- rede: interfaces, MAC, IP, gateway, DNS, DHCP, Winsock, TCP/IP, Wi-Fi, ARP e conexões TCP;
- catálogo Winget;
- GLPI Agent;
- reinício/desligamento com confirmação;
- ações em lote;
- auditoria JSONL.

### Network / NAC Intelbras

- parser tolerante para formatos comuns de tabela MAC;
- normalização de MAC e OUI;
- classificação por fabricante homologado;
- lista de exceções por MAC exato;
- identificação de MAC localmente administrado/randomizado;
- relatório por porta/VLAN;
- auditoria por arquivo ou SSH;
- perfis Intelbras por família;
- geração de plano de MAC ACL por OUI para Série 3000;
- recusa automática de política OUI em perfil cuja sintaxe não esteja validada;
- exclusão explícita de uplinks/trunks/APs na geração;
- geração de plano sem aplicação automática;
- testes offline para parser, OUI e ACL.

## Estrutura

```text
central_n2/
├── main.py
├── network_nac_cli.py
├── pyproject.toml
├── config/
│   ├── settings.json
│   ├── oui_allowlist.json
│   └── oui_allowlist.example.json
├── core/
│   ├── executor.py
│   ├── logger.py
│   └── result.py
├── modules/
│   ├── batch.py
│   ├── diagnostics.py
│   ├── glpi.py
│   ├── network.py
│   ├── network_nac.py
│   ├── software.py
│   └── system.py
├── ui/
│   └── console.py
├── docs/
│   └── network_nac_intelbras.md
├── tests/
│   ├── test_core.py
│   └── test_network_nac.py
├── logs/
└── reports/
```

## Requisitos

- Windows 10/11 ou Windows Server na estação administrativa;
- Python 3.10+;
- privilégios administrativos para o módulo de endpoints;
- conectividade e permissões administrativas nos computadores alvo;
- WinRM/PowerShell Remoting para o caminho preferencial;
- PsExec para fallback quando necessário;
- cliente OpenSSH do Windows para auditoria direta de switches.

A aplicação usa somente a biblioteca padrão do Python em produção. `pytest` é dependência opcional de desenvolvimento.

## Execução da Central N2

```powershell
python .\central_n2\main.py
```

## Execução do Network / NAC

```powershell
cd central_n2
python .\network_nac_cli.py
```

O módulo NAC nasce em modo seguro: audita e gera plano, mas não salva nem aplica a ACL automaticamente.

## Configuração de endpoints

Edite:

```text
central_n2/config/settings.json
```

Os principais itens são timeout, caminho do PsExec, instalador GLPI e catálogo Winget.

## Configuração de OUI

Edite:

```text
central_n2/config/oui_allowlist.json
```

Estrutura:

```json
{
  "manufacturers": [
    {"name": "HP", "prefixes": ["00:11:22"]},
    {"name": "Epson", "prefixes": ["AA:BB:CC"]}
  ],
  "exact_macs": [
    {"name": "Exceção documentada", "mac": "DE:AD:BE:EF:00:01"}
  ]
}
```

Os valores do exemplo acima são apenas ilustrativos. Cadastre somente OUIs realmente validados para o seu parque.

## Segurança operacional do NAC

Nunca trate uma porta de infraestrutura como se fosse porta de usuário. Não aplique política de endpoint indiscriminadamente em:

- uplinks;
- trunks;
- outro switch;
- Access Point que transporte MACs de clientes;
- hypervisor;
- firewall/roteador;
- telefone IP com computador atrás;
- bridge;
- servidor que apresente múltiplos MACs.

O fluxo recomendado é:

```text
AUDITORIA
   ↓
INVENTÁRIO / CLASSIFICAÇÃO
   ↓
SANEAMENTO
   ↓
BANCADA EM 1 PORTA
   ↓
PILOTO
   ↓
EXPANSÃO
```

A documentação detalhada está em:

```text
central_n2/docs/network_nac_intelbras.md
```

## Perfis Intelbras atuais

### Série 3000

É o perfil habilitado para geração automática de plano OUI `/24`, usando MAC ACL padrão e wildcard.

### S2050G-A

Pode ser auditado e possui suporte de código para ACL exata. A geração automática de política OUI fica bloqueada até validação de sintaxe/firmware específica.

### S-Series

Pode ser auditado e possui suporte de código para ACL exata. A geração OUI automática também fica bloqueada nesta versão.

## WinRM e fallback

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

## Auditoria

Endpoints geram registros em:

```text
central_n2/logs/YYYY-MM-DD.jsonl
```

O Network/NAC gera relatórios locais em:

```text
central_n2/reports/network_nac/
```

A pasta `reports/` é ignorada pelo Git para evitar versionar informações reais da rede.

Não armazene senhas, tokens ou chaves privadas no repositório.

## Testes

```powershell
cd central_n2
python -m pip install pytest
python -m pytest -q
```

Os testes cobrem componentes centrais e também:

- normalização de MAC/OUI;
- conversão Intelbras;
- wildcard Série 3000;
- MAC localmente administrado;
- parser de tabela MAC;
- deduplicação;
- classificação autorizado/não autorizado;
- resumo por porta;
- recusa de ACL vazia;
- validação da faixa de ACL;
- recusa de política OUI em perfil não validado.

O workflow `.github/workflows/central-n2-tests.yml` compila o Python e executa `pytest` em Windows.

## Limitações conhecidas

- Winget depende de um contexto remoto em que o executável esteja funcional;
- ambientes sem WinRM dependem de `ADMIN$`, RPC/SMB e PsExec permitido;
- o instalador GLPI deve ser ajustado ao ambiente antes de uso externo;
- OUI reduz equipamentos indevidos, mas não é autenticação forte: MAC pode ser spoofado;
- o parser de tabela MAC foi feito para formatos comuns e deve ser validado com saída real de cada modelo/firmware;
- o módulo NAC não salva configuração de switch automaticamente;
- 802.1X/RADIUS continua sendo a evolução indicada para autenticação forte de endpoints.

## Princípios do projeto

1. preservar o código histórico;
2. separar interface, regra e transporte;
3. evitar parsing frágil quando dados estruturados forem possíveis;
4. preferir PowerShell/CIM ao WMIC;
5. manter PsExec como fallback;
6. registrar ações administrativas;
7. confirmar operações disruptivas;
8. manter auditoria de rede separada da administração de endpoints;
9. recusar automaticamente mudanças potencialmente destrutivas quando faltarem dados;
10. validar em bancada antes de produção.
