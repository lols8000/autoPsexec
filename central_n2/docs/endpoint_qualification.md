# Homologação real de endpoints — Central N2 5.2

## Objetivo

Esta etapa valida a Central N2 fora do CI, usando estações Windows reais e os mesmos transportes, módulos, `ExecutionPolicy`, `ExecutionEngine`, recovery e validators utilizados pelo produto.

O runner de homologação foi desenhado para **não executar mutações por padrão**. A matriz básica é somente-leitura. Ações reais só acontecem quando são explicitamente informadas com `--qualification-action`. Ações que podem derrubar conectividade/reiniciar a estação exigem autorização adicional e confirmação nominal do host.

## Saída

Cada host gera dois artefatos em `reports/qualification`:

- `qualification-HOST-AAAAMMDD-HHMMSS.json` — evidência estruturada completa;
- `qualification-HOST-AAAAMMDD-HHMMSS.md` — resumo operacional legível.

Os estados possíveis são:

- `PASS` — evidência suficiente e resultado confirmado;
- `FAIL` — falha comprovada;
- `UNKNOWN` — não foi possível comprovar o estado final;
- `SKIP` — caso não aplicável ao endpoint.

Um endpoint **não é homologado** se houver qualquer `FAIL` ou `UNKNOWN`.

## Perfis

| Perfil | Uso |
|---|---|
| `auto` | detecta capacidades e pula o que não se aplica |
| `local` | valida execução no próprio computador administrativo |
| `winrm` | exige que o transporte selecionado seja WinRM |
| `psexec` | exige que o transporte selecionado seja PsExec |
| `notebook` | cenário de notebook/Wi-Fi; transporte é detectado |
| `printer` | estação de impressão; transporte é detectado |
| `glpi` | estação com GLPI Agent; transporte é detectado |
| `domain` | estação membro de domínio; transporte é detectado |

## Matriz automática somente-leitura

A execução básica coleta:

1. conectividade e transporte administrativo;
2. capabilities;
3. health snapshot;
4. processos;
5. serviços;
6. adaptadores;
7. IP/Gateway/DNS;
8. dispositivos PnP com problema;
9. postura de segurança;
10. domínio/secure channel/horário quando aplicável;
11. impressoras quando PrintManagement estiver disponível;
12. GLPI quando detectado;
13. Windows Update quando COM API estiver disponível.

Casos não aplicáveis viram `SKIP`, e não falso `FAIL`.

## Comandos recomendados

### 1. Endpoint local

```powershell
.\CentralN2.exe --qualify localhost --qualification-profile local
```

### 2. Endpoint WinRM

```powershell
.\CentralN2.exe --qualify HUMAP-WK-XXXXX --qualification-profile winrm
```

### 3. Endpoint PsExec

```powershell
.\CentralN2.exe --qualify HUMAP-WK-XXXXX --qualification-profile psexec
```

Este cenário é especialmente importante em redes onde TCP/5985 está bloqueado e ADMIN$/445 permanece permitido.

### 4. Vários endpoints, somente-leitura

```powershell
.\CentralN2.exe `
  --qualify PC-001 `
  --qualify PC-002 `
  --qualify PC-003 `
  --qualification-profile auto
```

Múltiplos hosts não podem ser usados quando ações disruptivas estiverem autorizadas.

## Ações reais de baixo risco

A ação precisa ser explicitamente listada. Exemplo:

```powershell
.\CentralN2.exe `
  --qualify PC-001 `
  --qualification-action network.flush_dns
```

Outras ações recomendadas para homologação controlada, apenas quando o endpoint for compatível:

```text
printer.restart_spooler
glpi.force_inventory
domain.gpupdate
defender.signatures
```

Todas passam pelo `ExecutionPolicy`. Capability, transporte ou privilégio incompatível bloqueiam a ação antes da mutação.

## Ações com parâmetros

Formato:

```text
ACTION:KEY=VALUE
```

Exemplo de restart de NIC:

```powershell
--qualification-param "network.adapter_restart:adapter_name=Ethernet"
```

## Ações disruptivas

Ações que podem derrubar transporte ou reiniciar a máquina exigem **três condições simultâneas**:

1. `--qualification-action ACTION`;
2. `--qualification-allow-disruptive`;
3. `--qualification-confirm-host HOST` exatamente igual ao host testado.

### DHCP renew/recovery

```powershell
.\CentralN2.exe `
  --qualify PC-001 `
  --qualification-action network.renew_dhcp `
  --qualification-allow-disruptive `
  --qualification-confirm-host PC-001
```

### Restart do adaptador

```powershell
.\CentralN2.exe `
  --qualify PC-001 `
  --qualification-action network.adapter_restart `
  --qualification-param "network.adapter_restart:adapter_name=Ethernet" `
  --qualification-allow-disruptive `
  --qualification-confirm-host PC-001
```

### Reboot/recovery

```powershell
.\CentralN2.exe `
  --qualify PC-001 `
  --qualification-action energy.restart `
  --qualification-param "energy.restart:delay_seconds=0" `
  --qualification-allow-disruptive `
  --qualification-confirm-host PC-001
```

O reboot usa o mesmo fluxo de recovery do ExecutionEngine: queda esperada, tentativa de reabertura da sessão administrativa e postcheck após recuperação.

## Matriz de campo recomendada

| ID | Cenário | Perfil | Critério principal |
|---|---|---|---|
| Q01 | computador administrativo | `local` | `READY_LOCAL` |
| Q02 | domínio + WinRM | `winrm` | `READY_WINRM` |
| Q03 | WinRM indisponível + ADMIN$/PsExec | `psexec` | `READY_PSEXEC` |
| Q04 | notebook/Wi-Fi | `notebook` | baseline + recovery controlado |
| Q05 | estação com impressora | `printer` | PrintManagement/Spooler |
| Q06 | estação com GLPI | `glpi` | GLPI status/inventory |
| Q07 | domínio | `domain` | secure channel/GPO/horário |
| Q08 | capability deliberadamente ausente | `auto` | ação incompatível é bloqueada pela policy |

## Ordem segura de execução em campo

1. rode a matriz somente-leitura;
2. revise `FAIL` e `UNKNOWN`;
3. só depois execute ações de baixo risco;
4. valide evidências;
5. reserve DHCP/NIC/reboot para endpoint de teste controlado;
6. nunca autorize uma ação disruptiva em múltiplos hosts;
7. mantenha o JSON de evidência junto do chamado/registro de homologação.

## Critério de conclusão da Etapa 2

A Etapa 2 só pode ser marcada como concluída quando existir evidência real para os perfis disponíveis no ambiente e, no mínimo:

- Q01 aprovado;
- Q02 ou Q03 aprovado em uma estação remota real;
- recovery de pelo menos uma ação de rede ou reboot comprovado em endpoint controlado;
- zero P0/P1 conhecidos;
- todo `FAIL`/`UNKNOWN` analisado e classificado;
- artefatos JSON/Markdown preservados.

CI e mocks **não substituem** esta etapa. Eles apenas provam que o runner e seus controles funcionam; a homologação exige endpoints reais.
