# Configuração — Central N2 Workstation v3

## 1. Arquivo principal

A configuração atualmente é carregada de:

```text
central_n2/config/settings.json
```

Estrutura atual:

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

---

## 2. `timeout_seconds`

```json
"timeout_seconds": 60
```

Timeout padrão usado pelo `RemoteExecutor` em operações comuns.

### Recomendações

- 30–60 s: consultas simples em LAN estável;
- 60–120 s: ambientes mais lentos ou com estações em redes remotas;
- não aumentar indiscriminadamente para “resolver” falhas de conectividade;
- operações longas possuem timeouts próprios na UI/módulos.

Um timeout recorrente deve ser investigado como sintoma de DNS, rede, WinRM, firewall, serviço remoto ou performance da estação.

---

## 3. `psexec_path`

```json
"psexec_path": "C:\\Windows\\System32\\PsExec.exe"
```

Caminho preferencial do `PsExec.exe`.

### Validação

```powershell
Test-Path "C:\Windows\System32\PsExec.exe"
```

Se o PsExec estiver em outro local, ajuste o caminho.

### Segurança

- use somente binário confiável;
- mantenha controle da origem e versão;
- não inclua credenciais em linha de comando persistida/configuração;
- avalie políticas internas antes de distribuir PsExec.

---

## 4. `sysinternals_dir`

```json
"sysinternals_dir": "C:\\Sysinternals"
```

Diretório local na estação administrativa onde ficam ferramentas opcionais da suíte Sysinternals.

Integrações atuais incluem:

- `autorunsc.exe`;
- `procdump.exe`;
- `handle.exe`;
- `sigcheck.exe`.

A aplicação pode verificar a existência de outras ferramentas, mas não as baixa automaticamente.

### Exemplo de estrutura

```text
C:\Sysinternals\
├── Autorunsc.exe
├── ProcDump.exe
├── Handle.exe
├── Sigcheck.exe
├── PsPing.exe
├── PsLoggedOn.exe
├── RAMMap.exe
└── Procmon.exe
```

---

## 5. `glpi`

```json
"glpi": {
  "installer_source": "",
  "remote_installer_path": "C:\\glpiagentinstall.vbs",
  "service_names": ["glpi-agent", "GLPI-Agent"]
}
```

### `installer_source`

Origem do instalador GLPI utilizado pelo ambiente.

O repositório público mantém esse campo vazio de propósito.

Exemplo local:

```json
"installer_source": "\\\\servidor\\share\\glpiagentinstall.vbs"
```

**Não publique** caminho interno real em repositório público sem necessidade.

### `remote_installer_path`

Caminho onde o instalador será colocado/executado na estação alvo.

### `service_names`

Lista de nomes de serviço que a Central considera como possíveis nomes do GLPI Agent.

---

## 6. `software`

Catálogo de aplicações homologadas para operações Winget.

Formato:

```json
"software": {
  "chave": {
    "name": "Nome amigável",
    "winget_id": "Publisher.Package"
  }
}
```

Exemplo:

```json
"chrome": {
  "name": "Google Chrome",
  "winget_id": "Google.Chrome"
}
```

### Boas práticas

- usar IDs exatos;
- validar instalação em uma estação de teste;
- não presumir que Winget funciona no mesmo contexto sob PsExec/SYSTEM;
- documentar aplicações corporativas específicas;
- evitar comandos de instalação arbitrários inseridos pelo operador.

---

## 7. `compliance`

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

### `min_disk_free_percent`

Percentual mínimo de espaço livre esperado no disco C:.

### `max_uptime_days`

Uptime acima desse valor gera desvio/penalidade conforme a lógica atual.

### `defender_required`

Define se Defender desabilitado é considerado não conformidade.

### `firewall_required`

Define se ausência de perfil de Firewall habilitado é não conformidade.

### `glpi_required`

Define se GLPI Agent não executando gera desvio.

### `pending_reboot_not_allowed`

Define se reboot pendente é considerado não conformidade.

### Observação importante

Compliance é um baseline operacional da Central, não substitui política formal de segurança, GPO, CIS Benchmark, EDR ou auditoria de conformidade corporativa.

---

## 8. `ui`

```json
"ui": {
  "heartbeat_seconds": 0.2,
  "long_operation_timeout_seconds": 3600
}
```

### `heartbeat_seconds`

Intervalo visual entre atualizações do spinner/tempo decorrido.

O runner aplica limite mínimo interno de 0,05 s.

Recomendação:

```text
0.1–0.5 s
```

Valores muito baixos aumentam escrita no console sem trazer benefício operacional.

### `long_operation_timeout_seconds`

Timeout padrão usado pela UI para operações longas quando o menu não define valor mais específico.

`3600` representa 1 hora.

### Importante

Timeout local não é mecanismo garantido de cancelamento remoto. Consulte [`architecture.md`](architecture.md).

---

## 9. Configuração por ambiente

A aplicação **ainda não implementa** merge automático com `settings.local.json`.

Portanto, hoje existem três alternativas:

1. editar `settings.json` somente na cópia operacional e nunca commitar valores internos;
2. manter patch/configuração privada fora do repositório público;
3. implementar futuramente um `SettingsLoader` com arquivo local ignorado pelo Git.

A terceira opção é a recomendação arquitetural para evolução.

---

## 10. O que nunca colocar no `settings.json` público

- senha de domínio;
- usuário/senha administrativa;
- token;
- API key;
- chave privada SSH;
- segredo de aplicação;
- credencial de proxy;
- senha de compartilhamento;
- caminhos internos desnecessários;
- dados pessoais;
- inventários reais de estações.

---

## 11. Validação antes de uso

Após editar configuração:

```powershell
python -m json.tool .\central_n2\config\settings.json
```

Depois execute os testes:

```powershell
cd central_n2
python -m pytest -q
```

E faça primeiro uma validação em estação de teste.

---

## 12. Exemplo de checklist de configuração

```text
[ ] Python >= 3.10
[ ] PsExec no caminho configurado
[ ] Sysinternals no diretório esperado, se utilizado
[ ] installer_source GLPI definido localmente, se necessário
[ ] catálogo Winget validado
[ ] baseline de compliance aprovado
[ ] timeout coerente com a rede
[ ] heartbeat visível e não excessivo
[ ] teste em estação de bancada concluído
[ ] nenhum segredo versionado
```
