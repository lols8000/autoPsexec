# Central N2 Workstation 5.1

> Plataforma de troubleshooting e remediação controlada para estações Windows, evoluída do projeto histórico **autoPsexec**.

A aplicação atual está em \`central_n2/\`. Os scripts antigos da raiz permanecem apenas como referência histórica.

## Fluxo operacional

\`\`\`text
EVIDÊNCIA
  ↓
DIAGNÓSTICO
  ↓
REMEDIAÇÃO
  ↓
VALIDAÇÃO
  ↓
REGISTRO
\`\`\`

## Destaques da 5.1

- Local, WinRM e PsExec como transportes explícitos;
- WinRM só é considerado pronto após \`Invoke-Command\` autenticado;
- fallback seguro: consultas podem repetir, mutações indeterminadas não;
- \`AttendanceContext\` como fonte única de estado do atendimento;
- scheduler compartilhado com serialização de mutações por host;
- health/compliance com PASS, FAIL, UNKNOWN e N/A;
- diagnóstico correlacionado e nove playbooks;
- remediações guiadas com before/after e validadores específicos;
- SQLite com migrations, retenção e correlation_id;
- relatórios Markdown/JSON/TXT e integração GLPI opcional;
- updater com SemVer, download atômico e SHA-256 quando disponível;
- CI Windows em Python 3.10/3.12/3.13, Ruff, mypy, coverage e build;
- PyInstaller + Inno Setup, hashes de release e Authenticode opcional.

## Executar

\`\`\`powershell
git clone https://github.com/lols8000/autoPsexec.git
cd autoPsexec
python .\central_n2\main.py
\`\`\`

GUI opcional:

\`\`\`powershell
python .\central_n2\main.py --gui
\`\`\`

## Transporte

\`\`\`text
LOCAL
  → execução direta

REMOTO
  → WinRM autenticado?
       → SIM: WinRM
       → NÃO: PsExec/ADMIN$ válidos?
            → SIM: PsExec
            → NÃO: diagnóstico de conectividade
\`\`\`

Em ambientes onde 5985 está bloqueada, PsExec pode manter o atendimento se TCP 445, ADMIN$ e privilégios administrativos estiverem disponíveis.

## Segurança de mutação

Se uma ação pode ter chegado ao host e a confirmação se perde, a Central marca o resultado como **indeterminado** e bloqueia fallback automático. O técnico valida o estado antes de repetir.

## Configuração

Pública:

\`\`\`text
central_n2\config\settings.json
\`\`\`

Local/não versionada:

\`\`\`text
central_n2\config\settings.local.json
\`\`\`

Nunca versione tokens, credenciais, URLs internas sensíveis, logs ou relatórios operacionais.

## Documentação

Comece por:

- \`central_n2/README.md\`;
- \`central_n2/docs/README.md\`;
- \`central_n2/docs/architecture.md\`;
- \`central_n2/docs/operations.md\`;
- \`central_n2/docs/security.md\`;
- \`central_n2/CHANGELOG.md\`.

## Qualidade

\`\`\`powershell
cd central_n2
python -m pip install -r requirements-dev.txt
python -W error::SyntaxWarning -m compileall -q .
python -m pytest -q
\`\`\`

O CI aplica gates adicionais de Ruff, mypy, coverage, matriz Python e build de distribuição.
