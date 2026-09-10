# Release e distribuição — Central N2 Workstation 5.1.0

## Fontes de versão

A versão deve permanecer sincronizada em:

- \`core/version.py\`;
- \`VERSION\`;
- \`pyproject.toml\`;
- \`version_info.txt\`;
- \`installer/CentralN2.iss\`.

Há teste automatizado que falha quando esses valores divergem.

## Dependências de build

\`requirements-dev.txt\` fixa versões de:

- pytest;
- pytest-cov;
- PyInstaller;
- Ruff;
- mypy.

O runtime da Central permanece baseado majoritariamente na biblioteca padrão.

## Gate de CI

Em pull request/push relevante:

### Matriz de testes

Windows com Python:

- 3.10;
- 3.12;
- 3.13.

### Gate 3.12

- compile com SyntaxWarning tratado como erro;
- Ruff para erros de correção/import;
- mypy nos contratos endurecidos;
- pytest com coverage e piso mínimo.

### Build

Após a matriz:

- PyInstaller onedir;
- Inno Setup 6.7.1;
- smoke do instalador.

## Reprodutibilidade

GitHub Actions são referenciadas por SHA imutável. Inno Setup e dependências Python de CI são fixados.

## PyInstaller

\`CentralN2.spec\`:

- inclui \`settings.json\` e baselines;
- inclui metadata de versão Windows;
- usa \`upx=False\`;
- gera pacote \`onedir\`.

\`settings.local.json\` nunca deve entrar no artefato.

## Instalador

\`installer/CentralN2.iss\` gera \`CentralN2-Setup.exe\`.

O instalador exige privilégio administrativo e instala a aplicação em arquitetura x64 compatível.

## Assinatura Authenticode

O workflow suporta assinatura opcional do EXE portátil e do instalador quando os secrets de certificado estão configurados.

Sem certificado, o release continua possível, mas o artefato permanece unsigned. A ausência de assinatura deve ser uma decisão consciente do mantenedor.

## Hashes

O workflow gera \`SHA256SUMS.txt\` para:

- ZIP portátil;
- instalador.

O manifesto é publicado no release.

## Atualizador interno

O updater:

1. consulta a release mais recente;
2. compara versões por SemVer;
3. valida nome do asset;
4. exige URL HTTPS;
5. baixa para arquivo temporário;
6. valida tamanho publicado;
7. valida SHA-256 quando \`digest=sha256:...\` estiver disponível;
8. executa \`fsync\`;
9. promove por substituição atômica.

A Central não faz auto-replace silencioso do executável em uso.

## Fluxo de promoção

\`\`\`text
branch
  ↓
CI verde
  ↓
PR
  ↓
CI do PR verde
  ↓
merge com head SHA esperado
  ↓
tag v5.1.0
  ↓
workflow de release
  ↓
artefatos + SHA256SUMS + assinatura opcional
\`\`\`

## Comandos locais equivalentes

\`\`\`powershell
cd central_n2
python -m pip install -r requirements-dev.txt
python -W error::SyntaxWarning -m compileall -q .
python -m ruff check .
python -m pytest -q
pyinstaller CentralN2.spec --clean --noconfirm
\`\`\`

O conjunto exato de paths/gates do CI é definido nos workflows versionados.

## Política de release

Não publicar tag se:

- versão estiver divergente;
- CI do PR estiver vermelho;
- houver segredo/configuração interna no diff;
- updater/build/instalador não tiverem smoke verde;
- documentação não refletir o comportamento real.
