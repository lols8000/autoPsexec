# Central de Manutenção N2 — Estações Windows

Evolução do antigo `autoPsexec` para uma central modular de diagnóstico, manutenção, conformidade e administração remota de **estações de trabalho Windows**.

A v2 remove do escopo operações de switch, NAC, OUI e MAC ACL. O projeto passa a ser exclusivamente orientado a endpoints Windows.

## O que está implementado

### Visão geral / Saúde

- pré-flight com DNS, ping, `ADMIN$` e WinRM;
- fabricante, modelo, serial, Windows e usuário;
- CPU, RAM, disco e uptime;
- reinicialização pendente;
- serviços automáticos parados;
- Defender, Firewall e GLPI Agent;
- score de saúde com achados por severidade.

### Usuários e perfis

- sessões abertas;
- administradores locais;
- perfis locais;
- último uso;
- tamanho aproximado do perfil;
- remoção por SID com confirmação.

### Rede da estação

- interfaces;
- IP, gateway e DNS;
- DHCP;
- flush DNS;
- reset Winsock/TCP-IP;
- Wi-Fi;
- ARP;
- conexões TCP;
- teste de porta remoto.

### Software

- inventário de programas;
- Winget;
- catálogo homologado;
- instalação, atualização e remoção controladas.

### Processos e serviços

- processos por consumo;
- finalização controlada;
- serviços;
- iniciar, parar e reiniciar.

### Windows Update

- histórico de hotfixes;
- atualizações pendentes;
- disparo de busca;
- reset controlado de BITS/Windows Update/Cryptographic Services.

### Segurança

- Microsoft Defender;
- proteção em tempo real;
- assinatura e varreduras;
- ameaças recentes;
- Firewall;
- BitLocker;
- TPM;
- Secure Boot;
- RDP;
- SMBv1;
- UAC.

A Central não oferece ações para desativar Defender ou Firewall.

### Domínio / GPO

- domínio e participação no domínio;
- secure channel;
- DC via `nltest`;
- hora via `w32tm`;
- `gpresult`;
- `gpupdate /force`;
- reparo controlado do secure channel.

### Impressoras

- impressoras instaladas;
- drivers e portas;
- filas;
- reinício do Spooler;
- limpeza de filas com confirmação.

### GLPI Agent

- status;
- instalação/reparo;
- serviço;
- inventário forçado;
- log recente.

### Disco / limpeza

- uso do disco C:;
- tamanho dos perfis;
- estimativa de espaço temporário recuperável;
- limpeza segura de TEMP e Lixeira.

Downloads não são apagados automaticamente.

### Eventos e diagnóstico

- eventos críticos recentes;
- pacote de diagnóstico local em JSON;
- informações de SO, hardware, rede, serviços, eventos, impressoras e processos.

### Compliance

Baseline configurável em `config/settings.json`:

```json
"compliance": {
  "min_disk_free_percent": 15,
  "max_uptime_days": 30,
  "defender_required": true,
  "firewall_required": true,
  "glpi_required": true,
  "pending_reboot_not_allowed": true
}
```

A Central calcula um score e mostra cada controle aprovado ou reprovado.

### Ações em lote

- ping;
- GPUpdate;
- flush DNS;
- inventário GLPI;
- snapshot de saúde.

## Estrutura

```text
central_n2/
├── main.py
├── config/
│   └── settings.json
├── core/
│   ├── executor.py
│   ├── logger.py
│   └── result.py
├── modules/
│   ├── batch.py
│   ├── compliance.py
│   ├── diagnostic_package.py
│   ├── diagnostics.py
│   ├── disk.py
│   ├── domain.py
│   ├── glpi.py
│   ├── health.py
│   ├── network.py
│   ├── printers.py
│   ├── security.py
│   ├── software.py
│   ├── system.py
│   ├── updates.py
│   └── users_profiles.py
├── ui/
│   └── console.py
├── docs/
│   └── workstation_v2.md
├── tests/
│   ├── test_core.py
│   └── test_workstation_v2.py
├── logs/
└── reports/
```

## Requisitos

- Windows 10/11 ou Windows Server na estação administrativa;
- Python 3.10+;
- privilégios administrativos;
- conectividade/permissões administrativas nas estações alvo;
- WinRM/PowerShell Remoting preferencial;
- PsExec como fallback.

A aplicação utiliza apenas biblioteca padrão do Python em produção. `pytest` é dependência opcional para desenvolvimento.

## Execução

```powershell
python .\central_n2\main.py
```

Se não estiver elevado, o programa solicita UAC automaticamente.

## Testes

```powershell
cd central_n2
python -m pip install pytest
python -m pytest -q
```

O workflow do GitHub Actions compila todos os arquivos Python e executa a suíte em Windows.

## Segurança operacional

Operações disruptivas exigem confirmação textual. Não armazene senhas, tokens, chaves ou credenciais no repositório.

A ferramenta foi projetada para uso administrativo autorizado e deve ser validada primeiro em estações de teste/piloto.

## Documentação detalhada

Consulte:

```text
central_n2/docs/workstation_v2.md
```
