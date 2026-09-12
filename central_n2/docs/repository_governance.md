# Governança do repositório — proteção do master

A Central N2 5.2 mantém a política do branch `master` versionada em
`.github/rulesets/master-protection.json`.

## Política adotada

O `master` deve aceitar mudanças somente por Pull Request, exigir branch
atualizada e os quatro checks abaixo verdes antes do merge:

- `test (3.10)`
- `test (3.12)`
- `test (3.13)`
- `build`

O ruleset também bloqueia exclusão do `master`, bloqueia force push e exige
resolução das conversas de review.

O número de aprovações obrigatórias permanece em **0** enquanto houver apenas
um mantenedor efetivo. Exigir aprovação do próprio autor criaria um bloqueio
operacional sem ganho real. Quando houver um segundo mantenedor ativo, subir
`required_approving_review_count` para **1**.

## Aplicação

Execute em PowerShell autenticado como administrador do repositório:

~~~powershell
gh auth login
.\scripts\configure_master_protection.ps1
~~~

Para validar apenas o arquivo declarativo, sem chamar a API:

~~~powershell
.\scripts\configure_master_protection.ps1 -ValidateOnly
~~~

A operação é idempotente: se o ruleset `Protect master - Central N2` já
existir, ele é atualizado; caso contrário, é criado.

## Permissão necessária

A identidade usada pelo GitHub CLI precisa de permissão de administração
(`Administration: write`) no repositório.
