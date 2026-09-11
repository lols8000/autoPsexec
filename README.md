# Central N2 Workstation 5.2

> Plataforma de diagnóstico, execução remota controlada, validação e auditoria para estações Windows.

A aplicação atual está em `central_n2/`. Os scripts históricos da raiz permanecem apenas como referência.

## Fluxo operacional

```text
EVIDÊNCIA
  ↓
DIAGNÓSTICO
  ↓
PLANO / PRECONDITIONS
  ↓
EXECUÇÃO CONTROLADA
  ↓
POSTCHECK / RECOVERY
  ↓
PASS | FAIL | UNKNOWN
  ↓
AUDITORIA / ROLLBACK QUANDO POSSÍVEL
```

## Destaques da 5.2

- menu **28 - Central de Execuções** com catálogo dinâmico;
- catálogos separados por domínio em `execution/catalogs/`;
- `ExecutionAction` com risco, timeout, idempotência, retry, transportes, privilege, capabilities, disconnect mode, rollback e tags;
- `ExecutionPolicy` bloqueia ação incompatível antes da confirmação;
- seletores contextuais para processos, serviços, adaptadores, impressoras, perfis, PnP e sessões;
- busca de ações por nome, descrição e tags;
- confirmação reforçada para ações HIGH/CRITICAL, destrutivas ou disruptivas;
- perda de conexão esperada tratada como workflow, não erro genérico;
- rollback real para ações reversíveis, como estado/StartType de serviço, Registro homologado e move/rename protegido;
- pacotes, certificados e Registro somente por allowlist validada no bootstrap;
- SQLite schema v4 com histórico rico de execuções e rollback;
- audit payload sanitizado, parâmetros sensíveis redigidos e correlation_id;
- Local, WinRM e PsExec com fallback seguro e sem dupla execução cega;
- CI Windows em Python 3.10/3.12/3.13, Ruff, mypy, coverage, PyInstaller e Inno Setup.

## Executar

```powershell
git clone https://github.com/lols8000/autoPsexec.git
cd autoPsexec
python .\central_n2\main.py
```

GUI opcional:

```powershell
python .\central_n2\main.py --gui
```

## Transporte

```text
LOCAL
  → execução direta

REMOTO
  → WinRM autenticado?
       → SIM: WinRM
       → NÃO: PsExec/ADMIN$ válidos?
            → SIM: PsExec
            → NÃO: diagnóstico de conectividade
```

`READY_WINRM` exige execução autenticada de `Invoke-Command`, não apenas listener WSMan.

## Segurança de mutação

Consultas podem fazer fallback. Mutações seguem regra mais rígida:

```text
falha comprovadamente pré-execução
  → fallback/retry permitido conforme contrato

entrega incerta ou timeout remoto
  → indeterminate=True
  → sem retry automático cego
  → validar estado antes de decidir
```

## Central de Execuções

O menu 28 organiza ações por domínio. Antes de executar, a Central:

1. coleta parâmetros;
2. resolve seletores quando possível;
3. avalia sessão, transporte, privilege e capabilities;
4. executa preconditions específicas;
5. mostra o plano;
6. solicita confirmação;
7. executa serializado por host;
8. trata disconnect/recovery quando esperado;
9. executa postcheck;
10. persiste evidência e oferece rollback quando suportado.

Não existe shell remoto livre, editor genérico de Registro ou instalador arbitrário.

## Configuração

Base pública:

`central_n2/config/settings.json`

Override local, não versionado:

`central_n2/config/settings.local.json`

Use `settings.local.example.json` como modelo. Segredos, caminhos internos e allowlists corporativas ficam no arquivo local.

## Documentação

Comece por:

- `central_n2/docs/README.md`
- `central_n2/docs/architecture.md`
- `central_n2/docs/operations.md`
- `central_n2/docs/configuration.md`
- `central_n2/docs/security.md`
- `central_n2/docs/development.md`

## Qualidade

```powershell
cd central_n2
python -m pip install -r requirements-dev.txt
python -W error::SyntaxWarning -m compileall -q .
python -m pytest -q
```

O CI adiciona Ruff, mypy, coverage, matriz Python e smoke de distribuição.
