# Segurança — Central N2 Workstation 5.1.0

## Modelo de confiança

A Central é ferramenta administrativa privilegiada para operação autorizada em estações Windows.

Pressupostos:

- operador autorizado;
- estação administrativa confiável;
- rede e hosts dentro do escopo permitido;
- credenciais tratadas pela política corporativa;
- PsExec/Sysinternals homologados quando utilizados.

## Elevação

\`main.py\` solicita UAC quando necessário. A Central não contorna a política de elevação do Windows.

## Segredos

Nunca versione:

- senhas;
- tokens;
- API keys;
- Authorization headers;
- certificados privados;
- credenciais de domínio;
- URLs internas sensíveis.

Use \`settings.local.json\`, que não deve ser commitado.

## Auditoria e redaction

O logger aplica redaction em nomes/valores sensíveis e é compacto por padrão.

\`logging.verbose_payloads=false\` evita persistir command/stdout/data completos em cada resultado. Habilite payload verboso somente quando houver necessidade operacional e proteção adequada do diretório de logs.

## WinRM

A Central não considera WinRM pronto apenas porque \`Test-WSMan\` respondeu. O preflight também valida \`Invoke-Command\`.

Evite:

- \`TrustedHosts=*\`;
- desabilitar Firewall/Defender para “fazer funcionar”;
- habilitar listener fora da política corporativa.

## PsExec

PsExec é suportado, mas possui impacto administrativo relevante:

- usa SMB/ADMIN$;
- pode criar serviço remoto temporário;
- pode ser bloqueado por EDR;
- contexto SYSTEM pode diferir do usuário interativo;
- aplicativos dependentes de perfil, como Winget, podem não estar disponíveis.

A Central não baixa PsExec automaticamente.

## Fallback e idempotência

A regra de segurança mais importante da 5.1.0:

**consulta pode repetir; mutação não pode repetir cegamente.**

Para mutações, fallback WinRM → PsExec só ocorre quando a falha é classificada como pré-execução. Quando a ação pode ter chegado ao destino, o resultado vira indeterminado e a repetição automática é bloqueada.

## Entrada do operador

Entradas interpoladas em PowerShell devem usar validadores/quoting centralizados. Não exponha shell remoto livre na UI.

## Winget

Operações do Winget verificam \`$LASTEXITCODE\`. Exit code não-zero deve aparecer como falha, mesmo que o PowerShell em si não tenha lançado exceção automaticamente.

## Remediação

Ações de escrita devem possuir:

- \`ActionSpec\`;
- classe de operação;
- confirmação quando aplicável;
- timeout;
- resultado auditável;
- validador pós-ação quando o estado final for verificável.

## Concorrência

O JobManager serializa operações não somente-leitura por host, evitando sobreposição como DISM + cleanup + reboot na mesma estação.

## Persistência

SQLite, relatórios e logs podem conter dados de infraestrutura. Proteja com ACL e política de retenção.

O banco passa por \`quick_check\` na inicialização e usa migrations versionadas; não edite o schema manualmente em produção.

## GLPI API

Tokens ficam apenas em configuração local. A API é desabilitada por padrão.

## Supply chain / distribuição

- dependências de CI fixadas;
- GitHub Actions fixadas por SHA;
- UPX desativado;
- release gera SHA256SUMS;
- updater valida SHA-256 quando o asset publica digest;
- download é promovido apenas após validação;
- assinatura Authenticode pode ser aplicada pelo CI sem armazenar certificado no repositório.

## Repositório público

Não commite logs reais, dumps, relatórios, banco SQLite, nomes de usuários, inventários internos ou arquivos de configuração local.

## Incidente

Em suspeita de uso indevido:

1. interrompa a operação;
2. preserve logs e banco;
3. identifique \`correlation_id\`;
4. identifique estação administrativa e alvos;
5. revise jobs/remediações;
6. revise credenciais;
7. acione o processo corporativo de segurança.
