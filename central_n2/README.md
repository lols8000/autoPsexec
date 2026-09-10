# Central N2 Workstation — Guia rápido 5.1.0

A Central N2 é uma plataforma de troubleshooting e remediação controlada para estações Windows. O fluxo operacional é:

\`\`\`text
EVIDÊNCIA → DIAGNÓSTICO → REMEDIAÇÃO → VALIDAÇÃO → REGISTRO
\`\`\`

## Executar

Da raiz do repositório:

\`\`\`powershell
python .\central_n2\main.py
\`\`\`

GUI opcional:

\`\`\`powershell
python .\central_n2\main.py --gui
\`\`\`

A aplicação solicita elevação UAC quando necessário.

## Transporte

A seleção é automática:

\`\`\`text
alvo local
  → Local

alvo remoto
  → WinRM autenticado e executável?
      → SIM: WinRM
      → NÃO: ADMIN$ + PsExec válidos?
          → SIM: PsExec
          → NÃO: sem transporte administrativo
\`\`\`

\`READY_WINRM\` só é emitido depois que o preflight valida o listener e uma execução real de \`Invoke-Command\`.

Estados de conectividade:

- \`READY_LOCAL\`
- \`READY_WINRM\`
- \`READY_PSEXEC\`
- \`DNS_FAILED\`
- \`AUTHENTICATION_FAILED\`
- \`NETWORK_UNREACHABLE\`
- \`NO_USABLE_TRANSPORT\`

WinRM não é requisito absoluto. Em ambiente com 5985 bloqueada, a Central pode operar via PsExec quando TCP 445, ADMIN$ e privilégios administrativos estiverem disponíveis e o executável PsExec estiver homologado na estação administrativa.

## Fallback seguro

Leituras podem ser repetidas automaticamente em PsExec após falha de WinRM.

Mutações seguem outra regra:

\`\`\`text
falha comprovadamente antes da execução
  → fallback permitido

ação pode ter chegado ao host, mas a confirmação se perdeu
  → resultado INDETERMINADO
  → fallback automático bloqueado
  → validar o estado antes de repetir
\`\`\`

Isso evita executar duas vezes ações como reset, instalação, cleanup ou alteração de serviço.

## Compliance

A avaliação usa quatro estados:

- **PASS** — evidência disponível e dentro do baseline;
- **FAIL** — evidência disponível e fora do baseline;
- **UNKNOWN** — métrica não pôde ser obtida;
- **N/A** — controle não é exigido pelo baseline.

\`UNKNOWN\` não é tratado como \`FAIL\`. N/A não entra no denominador do score de compliance.

## Remediações guiadas

Atualmente possuem validação específica:

- limpeza segura de temporários;
- reinício do Spooler;
- reset de componentes do Windows Update;
- GPUpdate /force.

A Central coleta evidência antes/depois quando aplicável e registra o estado de validação como \`PASS\`, \`FAIL\` ou \`UNKNOWN\`.

## Configuração

Base pública:

\`\`\`text
config\settings.json
\`\`\`

Override local não versionado:

\`\`\`text
config\settings.local.json
\`\`\`

Use \`config\settings.local.example.json\` como modelo. O merge é recursivo. Segredos, URLs internas e caminhos privados ficam apenas no arquivo local.

## PsExec

Diretório recomendado:

\`\`\`text
C:\Sysinternals\PsExec.exe
\`\`\`

A Central também procura no PATH e em \`C:\Windows\System32\PsExec.exe\`.

Validação manual:

\`\`\`powershell
Test-NetConnection PC023 -Port 445
Test-Path \\PC023\ADMIN$
C:\Sysinternals\PsExec.exe -accepteula -nobanner \\PC023 cmd.exe /d /c echo CENTRAL_N2_OK
\`\`\`

## Persistência

Por padrão:

\`\`\`text
data\central_n2.db
\`\`\`

SQLite usa WAL, migrations versionadas e retenção configurável. O schema é validado na inicialização. Jobs, snapshots, findings, remediações e relatórios carregam \`correlation_id\`.

## Atualizações

O menu de atualização respeita \`updates.enabled\`. Downloads são feitos de forma controlada; a Central não se substitui silenciosamente.

Quando um asset publica digest SHA-256, o updater verifica o hash antes de promover o arquivo temporário ao destino final.

## Testes

\`\`\`powershell
cd central_n2
python -m pip install -r requirements-dev.txt
python -W error::SyntaxWarning -m compileall -q .
python -m pytest -q
\`\`\`

O CI também executa Ruff, mypy, coverage, matriz Python e smoke de build/instalador.

## Documentação

Consulte \`docs/README.md\` para arquitetura, operação, segurança, configuração, troubleshooting, desenvolvimento e distribuição.
