# Central N2 Workstation — Guia rápido 5.2.0

A Central N2 evoluiu de ferramenta de diagnóstico para uma plataforma operacional de suporte:

```text
EVIDÊNCIA → DIAGNÓSTICO → PLANO → EXECUÇÃO → VALIDAÇÃO → REGISTRO
```

## Iniciar

```powershell
python .\central_n2\main.py
```

A aplicação solicita UAC quando necessário.

## Selecionar estação

Use hostname/FQDN sempre que possível. O preflight produz uma sessão lógica com:

- conectividade;
- transporte selecionado;
- capabilities;
- contexto de execução;
- readiness administrativo.

Estados principais: `READY_LOCAL`, `READY_WINRM` e `READY_PSEXEC`.

## Menu 28 — Central de Execuções

O menu 28 é o ponto central para ações mutáveis. Há também atalhos `90+` nos menus de diagnóstico.

O fluxo é:

```text
ação
 ↓
parâmetros / seletor
 ↓
policy + capabilities + preconditions
 ↓
plano LIBERADO ou BLOQUEADO
 ↓
confirmação
 ↓
execução
 ↓
recovery se necessário
 ↓
postcheck
 ↓
PASS / FAIL / UNKNOWN
 ↓
SQLite / rollback quando disponível
```

### Confirmação

Ações comuns usam confirmação simples.

Ações HIGH/CRITICAL, destrutivas ou que podem afetar conectividade exigem:

```text
EXECUTAR <hostname>
```

Rollback disponível exige:

```text
DESFAZER <hostname>
```

## Seletores

Quando possível, a Central lista objetos reais da estação para evitar digitação manual:

- processos;
- serviços;
- adaptadores;
- impressoras;
- perfis;
- dispositivos PnP;
- sessões.

A opção manual permanece como fallback.

## Retry e resultado indeterminado

Cada ação possui `retry_policy`, `retry_attempts` e `retry_delay_seconds`.

Padrão: `PRE_EXECUTION_ONLY`.

A Central pode repetir apenas falha comprovadamente pré-execução (ou explicitamente `retry_safe` em ação idempotente). Mutação com entrega incerta nunca é repetida automaticamente. `indeterminate=True` implica validação do estado antes de nova tentativa.

## Desconexão esperada

Ações como DHCP renew, restart de NIC e reboot podem usar `DisconnectMode.TEMPORARY`.

A Central aguarda, reabre a sessão e só então roda o postcheck.

Shutdown usa `TERMINAL`: a perda de conectividade é consequência esperada e não dispara recovery de retorno.

## Rollback

Rollback existe apenas onde é tecnicamente defensável. Exemplos:

- Start/Stop de serviço → restaura estado anterior;
- StartType → restaura Automatic/Manual/Disabled anterior;
- Registro homologado → restaura valor/ausência anterior;
- move/rename protegido → move de volta à origem.

Não há rollback fictício para exclusão de arquivo, limpeza de TEMP ou remoção de perfil. Guardas específicas de rollback impedem reversões que poderiam sobrescrever estado novo.

## Catálogos corporativos

São validados no bootstrap:

- `packages`;
- `certificates`;
- `registry_actions`.

Entrada inválida é desabilitada e registrada; não derruba a Central inteira.

## Persistência

SQLite schema **4** registra execuções com:

- operador;
- versão da ação;
- transporte;
- início/fim/duração;
- risco;
- parâmetros redigidos;
- validation state;
- correlation_id;
- vínculo de rollback.

## Segurança

A Central não expõe:

- PowerShell livre;
- CMD livre;
- editor genérico de Registro;
- instalação arbitrária digitada pelo operador;
- importação de PFX/chave privada;
- delete recursivo genérico.

Consulte `docs/security.md`.

## Testes

```powershell
cd central_n2
python -W error::SyntaxWarning -m compileall -q .
python -m pytest -q
```

O CI oficial também executa Ruff, mypy, coverage, build portátil e instalador.
