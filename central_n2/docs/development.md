# Desenvolvimento — Central N2 Workstation 5.1.0

## Objetivo

Evoluir a Central sem degradar segurança, previsibilidade operacional ou manutenibilidade.

## Ambiente

- Windows;
- Python 3.10+;
- Git;
- dependências de desenvolvimento fixadas em \`requirements-dev.txt\`.

\`\`\`powershell
cd central_n2
python -m pip install -r requirements-dev.txt
\`\`\`

## Estrutura

\`\`\`text
central_n2/
├── main.py
├── config/
├── core/
│   ├── actions.py
│   ├── context.py
│   ├── evaluation.py
│   ├── executor.py
│   ├── jobs.py
│   ├── logger.py
│   └── transport/
├── diagnostics/
├── modules/
├── playbooks/
├── remediation/
├── reports/
├── storage/
├── ui/
│   ├── console_base.py
│   ├── console_v5.py
│   └── tk_app.py
├── tests/
└── docs/
\`\`\`

\`console_v3.py\` é histórico/compatibilidade. A UI operacional 5.1 não deve voltar a herdar essa classe.

## Dependência

\`\`\`text
UI
 ↓
orquestração/core
 ↓
módulos
 ↓
RemoteExecutor
 ↓
transportes
\`\`\`

Não permita:

- transporte importando UI;
- módulo chamando \`input()\`;
- regra de negócio específica dentro do executor;
- componente relendo configuração sem necessidade;
- mutação usando API de leitura apenas para obter fallback conveniente.

## Configuração

\`ConfigLoader\` é chamado no bootstrap. A configuração efetiva é injetada nos componentes principais.

Novo componente deve receber dependência/configuração pronta quando possível.

## Estado

Todo estado de um atendimento pertence a \`AttendanceContext\`.

Não crie aliases como:

\`\`\`python
self.last_report = self.context.report_path
\`\`\`

Use diretamente o contexto.

## ActionSpec

Operações de UI devem ser descritas por \`ActionSpec\` quando aplicável:

- chave estável;
- título;
- \`OperationClass\`;
- timeout;
- confirmação.

Nunca derive semântica de segurança a partir do texto exibido ao operador.

## OperationClass

- READ_ONLY;
- HEAVY_READ;
- LIGHT_WRITE;
- HEAVY_WRITE;
- DISRUPTIVE.

O JobManager serializa qualquer classe diferente de READ_ONLY por host.

## CommandResult

Módulos retornam \`CommandResult\`.

Dados estruturados vão em \`data\`; texto nativo relevante pode ficar em \`stdout\`.

Não retorne tuplas ad-hoc.

## Fallback

### Leitura

Pode usar \`execute_powershell\`, \`execute_powershell_json\` ou \`execute_cmd\` com fallback de leitura.

### Mutação

Use:

- \`execute_mutating_powershell\`;
- \`execute_mutating_powershell_json\`;
- \`execute_mutating_cmd\`.

Esses caminhos impedem dupla execução quando a entrega ao host é incerta.

## Resultado indeterminado

Ao adicionar mutação, trate \`CommandResult.indeterminate\`.

Não converta resultado indeterminado em sucesso só porque uma chamada local terminou sem exceção.

## PowerShell

Preferir objetos estruturados e \`ConvertTo-Json\`.

Evitar parsing textual quando existe cmdlet estruturado.

Entradas interpoladas devem usar validadores/quoting centralizados.

Para programas externos, verifique \`$LASTEXITCODE\` quando o exit code fizer parte do contrato.

## Remediação

Remediação deve separar:

1. probe anterior;
2. ação;
3. probe posterior;
4. validador;
5. persistência.

Validador deve devolver PASS, FAIL ou UNKNOWN.

UNKNOWN é obrigatório quando o estado final não pode ser provado.

## Avaliação

Use \`core/evaluation.py\` para compliance/health compartilhado.

Semântica:

- PASS: evidência confirma conformidade;
- FAIL: evidência confirma desvio;
- UNKNOWN: métrica ausente;
- NOT_APPLICABLE: controle desabilitado/não exigido.

Não use \`bool(None)\` para transformar métrica ausente em falha.

## Banco

Mudança de schema exige nova migration e incremento de \`SCHEMA_VERSION\`.

Não altere estrutura existente “in place” sem migration.

Adicione teste que cria banco antigo/novo quando a mudança afetar compatibilidade.

## Logging

Logs devem ser compactos por padrão. Não persista comandos/payloads completos sem necessidade.

Qualquer novo campo potencialmente sensível deve passar pelo mecanismo de redaction.

## Testes locais

\`\`\`powershell
python -W error::SyntaxWarning -m compileall -q .
python -m pytest -q
\`\`\`

Para reproduzir gates adicionais:

\`\`\`powershell
python -m ruff check .
python -m mypy --explicit-package-bases core/result.py core/context.py core/actions.py core/evaluation.py core/updater.py core/validation.py remediation/engine.py
\`\`\`

## CI

O workflow oficial executa:

- Python 3.10/3.12/3.13 em Windows;
- compile com SyntaxWarning como erro;
- Ruff;
- mypy em contratos críticos;
- pytest;
- coverage mínimo;
- PyInstaller;
- Inno Setup.

Não reduza um gate para “fazer o CI passar” quando o gate encontrou defeito real. Corrija o contrato ou justifique explicitamente a exceção.

## Checklist de PR

\`\`\`text
[ ] branch parte do master atual
[ ] entrada validada
[ ] operação classificada
[ ] mutação usa executor mutável
[ ] timeout definido
[ ] resultado indeterminado tratado
[ ] CommandResult preservado
[ ] dados estruturados quando possível
[ ] logging seguro
[ ] migration criada se schema mudou
[ ] testes adicionados
[ ] matriz/ruff/mypy/coverage verdes
[ ] build e instalador verdes
[ ] docs atualizadas
[ ] nenhum segredo no diff
\`\`\`

## Definition of Done

\`\`\`text
funciona
+ falha previsivelmente
+ não duplica mutação
+ mantém feedback
+ valida o resultado
+ deixa evidência
+ passa CI
+ está documentado
\`\`\`
