# Central N2 — Guia rápido v5

Versão: 5.0.0
Console: python main.py
GUI: python main.py --gui
Configuração local: copie config/settings.local.example.json para config/settings.local.json e nunca versione segredos.

A v5 detecta automaticamente quando o alvo é a própria máquina e usa transporte local, sem exigir WinRM. Destinos remotos usam WinRM preferencialmente e PsExec como fallback de transporte.

---

# Central N2 Workstation — Guia rápido

Este README é o guia rápido da aplicação localizada em `central_n2/`.

Para visão completa do projeto, consulte o [`README.md` da raiz](../README.md). Para detalhes técnicos, consulte [`docs/README.md`](docs/README.md).

---

## Versão operacional atual

**Central N2 Workstation v3 responsiva**

Características centrais:

- foco exclusivo em estações Windows;
- WinRM como transporte preferencial;
- PsExec como fallback;
- operações longas em worker threads;
- heartbeat visual e tempo decorrido;
- timeout padronizado;
- saúde/compliance;
- troubleshooting avançado;
- Sysinternals opcional;
- logging/auditoria local.

---

## Requisitos

- Windows 10/11 ou Windows Server na estação administrativa;
- Python 3.10+;
- privilégios administrativos;
- conectividade com estações alvo;
- WinRM/PowerShell Remoting preferencialmente configurado;
- PsExec disponível para fallback quando necessário;
- Sysinternals opcional;
- pytest somente para desenvolvimento/testes.

O runtime Python não possui dependências externas obrigatórias.

---

## Estrutura

```text
central_n2/
├── main.py
├── README.md
├── config/
│   └── settings.json
├── core/
│   ├── executor.py
│   ├── jobs.py
│   ├── logger.py
│   └── result.py
├── modules/
├── ui/
│   ├── console.py
│   └── console_v3.py
├── tests/
├── docs/
├── logs/
└── reports/
```

`ui/console_v3.py` é a interface atual. `ui/console.py` permanece como referência da geração anterior.

---

## Configuração inicial

Edite:

```text
config/settings.json
```

Configuração padrão atual:

```json
{
  "timeout_seconds": 60,
  "psexec_path": "C:\\Windows\\System32\\PsExec.exe",
  "sysinternals_dir": "C:\\Sysinternals",
  "glpi": {
    "installer_source": "",
    "remote_installer_path": "C:\\glpiagentinstall.vbs",
    "service_names": ["glpi-agent", "GLPI-Agent"]
  },
  "software": {
    "chrome": {"name": "Google Chrome", "winget_id": "Google.Chrome"},
    "firefox": {"name": "Mozilla Firefox", "winget_id": "Mozilla.Firefox"},
    "7zip": {"name": "7-Zip", "winget_id": "7zip.7zip"},
    "vnc": {"name": "UltraVNC", "winget_id": "uvncbvba.UltraVnc"}
  },
  "compliance": {
    "min_disk_free_percent": 15,
    "max_uptime_days": 30,
    "defender_required": true,
    "firewall_required": true,
    "glpi_required": true,
    "pending_reboot_not_allowed": true
  },
  "ui": {
    "heartbeat_seconds": 0.2,
    "long_operation_timeout_seconds": 3600
  }
}
```

O campo `glpi.installer_source` vem vazio porque o repositório é público. Configure o caminho real apenas no ambiente operacional e não publique credenciais ou informações internas.

Detalhes: [`docs/configuration.md`](docs/configuration.md).

---

## Iniciar

A partir da raiz do repositório:

```powershell
python .\central_n2\main.py
```

Ou dentro da pasta:

```powershell
cd central_n2
python .\main.py
```

A aplicação solicita elevação administrativa quando necessário.

---

## Menu atual

```text
╔════════════════════════════════════════════════════╗
║        CENTRAL N2 WORKSTATION — V3 RESPONSIVA     ║
╚════════════════════════════════════════════════════╝

Alvo: nenhum

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
[0] Sair
```

---

## Feedback visual

Operações bloqueantes são executadas por `ResponsiveJobRunner`.

Exemplo:

```text
| DISM RestoreHealth...   4.2s
/ DISM RestoreHealth...   8.7s
```

O spinner indica que a Central continua executando/aguardando. Não é porcentagem de progresso.

### Sucesso

```text
✓ SUCESSO [winrm] — 7524 ms
```

### Falha

```text
✗ FALHA [psexec]
Erro: Access denied
```

### Timeout

```text
✗ TIMEOUT
Operação 'DISM RestoreHealth' excedeu 3600s
```

**Atenção:** timeout local não garante interrupção de processo remoto já iniciado.

---

## Transportes

### Preferencial

```text
PowerShell Remoting / WinRM
```

### Fallback

```text
PsExec
```

Fluxo conceitual:

```text
Módulo
  ↓
RemoteExecutor
  ↓
WinRM disponível?
  ├─ sim → PowerShell Remoting
  └─ não → PsExec
```

---

## Principais módulos

| Área | Arquivo |
|---|---|
| Saúde | `modules/health.py` |
| Compliance | `modules/compliance.py` |
| Diagnóstico | `modules/diagnostics.py` |
| Performance | `modules/performance.py` |
| Reparo Windows | `modules/repair.py` |
| Drivers/dispositivos | `modules/devices.py` |
| Inicialização | `modules/startup.py` |
| Tarefas | `modules/tasks.py` |
| Crashes/BSOD | `modules/crashes.py` |
| Segurança | `modules/security.py` |
| Rede | `modules/network.py` |
| Usuários/perfis | `modules/users_profiles.py` |
| Software | `modules/software.py` |
| GLPI | `modules/glpi.py` |
| Impressoras | `modules/printers.py` |
| Domínio/GPO | `modules/domain.py` |
| Disco/limpeza | `modules/disk.py` |
| Storage/bateria | `modules/storage.py` |
| Ferramentas avançadas | `modules/workstation_tools.py` |
| Sysinternals | `modules/sysinternals.py` |
| Pacote diagnóstico | `modules/diagnostic_package.py` |
| Sistema | `modules/system.py` |
| Windows Update | `modules/updates.py` |

Descrição detalhada: [`docs/modules.md`](docs/modules.md).

---

## Sysinternals

Diretório padrão:

```text
C:\Sysinternals
```

Integrações atuais:

- Autorunsc;
- ProcDump;
- Handle;
- Sigcheck.

A Central não baixa essas ferramentas automaticamente.

---

## GLPI

Antes de usar instalação/reparo, defina localmente:

```json
"installer_source": "\\\\servidor\\share\\glpiagentinstall.vbs"
```

O valor não deve ser commitado no repositório público quando revelar infraestrutura interna.

---

## Logs

Logs administrativos:

```text
logs/YYYY-MM-DD.jsonl
```

Não versione logs reais.

---

## Relatórios

Pacotes de diagnóstico:

```text
reports/diagnostics/
```

Esses arquivos podem conter dados internos da estação e devem ser protegidos.

---

## Testes

```powershell
cd central_n2
python -m pip install pytest
python -m compileall .
python -m pytest -q
```

O CI executa compilação e pytest em Windows.

---

## Operação segura

Fluxo recomendado:

```text
Saúde
 ↓
Diagnóstico
 ↓
Confirmar causa
 ↓
Remediar
 ↓
Validar
 ↓
Gerar evidência
```

Evite usar SFC, DISM, reset de rede ou reboot como sequência automática para qualquer problema.

---

## Documentação completa

- [`docs/README.md`](docs/README.md) — índice
- [`docs/architecture.md`](docs/architecture.md) — arquitetura
- [`docs/modules.md`](docs/modules.md) — módulos
- [`docs/configuration.md`](docs/configuration.md) — configuração
- [`docs/operations.md`](docs/operations.md) — manual do suporte
- [`docs/security.md`](docs/security.md) — segurança
- [`docs/troubleshooting.md`](docs/troubleshooting.md) — troubleshooting
- [`docs/development.md`](docs/development.md) — desenvolvimento
