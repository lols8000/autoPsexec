# Homologação real de endpoints — Central N2 5.2

## Objetivo

A Etapa 2 do Release Readiness Plan valida a Central fora do CI, em endpoints Windows reais.

O processo é deliberadamente separado da operação normal da Central:

1. preflight e transporte;
2. capabilities;
3. health/read-only probes;
4. checks específicos por perfil;
5. ações controladas somente quando explicitamente autorizadas;
6. recovery e rollback usando o próprio ExecutionEngine;
7. evidência JSON + Markdown.

## Privacidade

Nunca commite a matriz real de endpoints.

Use o arquivo de exemplo apenas como modelo e salve a cópia real fora do Git, por exemplo:

    config/endpoint_matrix.local.json

Os relatórios ficam em:

    reports/field-validation/

Esse diretório já é ignorado pelo Git.

O relatório público não grava hostname/IP real. Ele contém:

- alias do perfil;
- fingerprint SHA-256 truncado;
- correlation ID;
- estado dos checks;
- metadata operacional limitada.

## Perfis mínimos de homologação

A campanha recomendada possui:

| Perfil | Objetivo |
|---|---|
| LOCAL-ADMIN | provar execução local/elevada |
| DOMAIN-WINRM | provar WinRM autenticado |
| DOMAIN-PSEXEC | provar ADMIN$ + fallback PsExec |
| NOTEBOOK-WIFI | provar NetAdapter/bateria |
| PRINT | provar PrintManagement/Spooler |
| GLPI | provar descoberta do agente |
| CONTROLLED-REBOOT | provar disconnect TEMPORARY + recovery |

Use máquinas de laboratório ou endpoints explicitamente autorizados para mutações.

## Execução SAFE

Abra PowerShell elevado no diretório central_n2.

    python -m validation.cli --matrix config/endpoint_matrix.local.json

SAFE não executa ações mutáveis, mesmo que elas estejam declaradas na matriz.

O exit code é:

- 0: PASS;
- 1: FAIL;
- 2: erro de configuração;
- 3: UNKNOWN.

## Execução CONTROLLED

Ações mutáveis exigem dois elementos:

1. a ação precisa estar declarada no endpoint da matriz;
2. a execução precisa receber consentimento explícito.

Exemplo:

    python -m validation.cli ^
      --matrix config/endpoint_matrix.local.json ^
      --allow-mutations ^
      --acknowledge "VALIDAR ENDPOINTS 5.2"

Ações HIGH exigem também:

    --allow-high-risk

O runner não automatiza ações destrutivas.

A única ação CRITICAL permitida pelo runner é:

    energy.restart

e exige:

    --allow-reboot

Shutdown, remoção de perfil, remoção de driver, delete arbitrário e disable destrutivo não fazem parte da campanha automatizada.

## Sequência recomendada

### Rodada A — SAFE

Execute todos os perfis habilitados.

Gate:

- todos os endpoints esperados devem estar READY;
- transporte deve corresponder ao perfil quando especificado;
- capabilities obrigatórias devem estar disponíveis;
- probes de health/rede e de cada role devem passar.

### Rodada B — Controlled LOW/MEDIUM

Sugestões:

- network.flush_dns;
- printer.restart_spooler;
- service.restart em serviço de laboratório explicitamente escolhido.

Cada ação é processada pelo ExecutionEngine normal:

    plan -> policy -> execute -> validate -> evidence

### Rodada C — HIGH

Somente em máquina de teste.

Use --allow-high-risk.

Valide especialmente:

- ação bloqueada quando capability falta;
- UNKNOWN quando postcheck não comprova estado;
- nenhum retry após indeterminate.

### Rodada D — reboot/recovery

Use apenas endpoint descartável/controlado.

Habilite energy.restart com --allow-reboot.

O resultado esperado é:

    comando entregue
      -> host indisponível
      -> recovery
      -> sessão READY novamente
      -> postcheck
      -> PASS

Se o host não retornar dentro da janela, o resultado deve permanecer UNKNOWN.

## Critérios de aprovação da Etapa 2

A etapa somente pode ser declarada concluída quando:

- LOCAL, WinRM e PsExec tiverem ao menos um endpoint real aprovado;
- nenhum P0/P1 estiver aberto;
- recovery tiver sido comprovado em pelo menos um reboot controlado;
- pelo menos uma ação reversível tiver rollback validado;
- policy tiver bloqueado corretamente ao menos um cenário incompatível;
- relatórios JSON/Markdown tiverem sido revisados sem dados sensíveis inesperados;
- todas as falhas reproduzíveis tiverem issue/regression test associado.

## Evidência a preservar

Para cada endpoint preserve localmente:

- relatório Markdown;
- JSON;
- correlation ID;
- versão da Central;
- data/hora;
- observação do técnico;
- ticket/issue quando houver falha.

Não publique logs corporativos reais em repositório público.
