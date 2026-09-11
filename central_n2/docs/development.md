# Desenvolvimento — Central N2 Workstation 5.2.0

## Objetivo

Adicionar capacidade sem degradar previsibilidade, segurança ou manutenção.

## Estrutura relevante

```text
central_n2/
├── core/
├── diagnostics/
├── execution/
│   ├── catalog.py
│   ├── engine.py
│   ├── models.py
│   ├── policy.py
│   ├── registry.py
│   ├── validators.py
│   ├── config_validation.py
│   └── catalogs/
├── modules/
├── remediation/
├── storage/
├── ui/
├── tests/
└── docs/
```

`execution/catalog.py` é apenas composition root. Não volte a concentrar dezenas de ações em um arquivo monolítico.

## Como adicionar uma ação

1. implemente a operação no módulo de domínio;
2. use API mutável do RemoteExecutor;
3. crie/edite o arquivo de domínio em `execution/catalogs/`;
4. descreva `ExecutionAction`;
5. adicione parâmetros tipados;
6. configure before/after probes;
7. use validator específico;
8. declare capabilities/transport/privilege;
9. declare disconnect/recovery quando aplicável;
10. declare rollback somente se real;
11. adicione tags;
12. escreva testes.

## ExecutionAction

Preencha conscientemente:

- key;
- title/category;
- OperationClass;
- RiskLevel;
- impact;
- timeout;
- confirmation/destructive/reboot;
- action_version;
- idempotent;
- retry_policy;
- retry_attempts;
- retry_delay_seconds;
- allowed_transports;
- required_capabilities;
- required_privilege;
- disconnect_mode;
- recovery timeout/delay;
- rollback_strategy;
- tags.

## Invariantes do registry

O registry rejeita:

- timeout <= 0;
- lista de transporte vazia/inválida;
- ação destrutiva sem confirmação;
- SAFE_TRANSIENT em ação não idempotente;
- TEMPORARY sem recovery timeout;
- rollback handler sem estratégia documentada.

Não desabilite essas validações para fazer catálogo carregar.

## Retry

O ExecutionEngine possui retry seletivo.

Regras:

- nunca repetir resultado `indeterminate`;
- `PRE_EXECUTION_ONLY` só repete resultado com `transport_failure_kind=pre_execution`;
- `SAFE_TRANSIENT` exige `idempotent=True` e pode aceitar `retry_safe=true`;
- respeitar `retry_attempts` e `retry_delay_seconds`.

O RemoteExecutor continua responsável por fallback de transporte e só faz fallback de mutação quando a falha é comprovadamente pré-execução.

Nunca implemente loop genérico de retry fora desse contrato.

## Preconditions

Use custom precondition quando a decisão depende do estado atual e não apenas de capability estática.

Precondition retorna PolicyCheck PASS/FAIL/WARN.

FAIL bloqueia handler.

## Disconnect

Use TEMPORARY para operações que derrubam conexão e devem voltar.

Use TERMINAL quando o efeito final esperado é perder a estação, como shutdown.

## Rollback

Rollback deve usar a evidência before original.

Requisitos:

- inversa real;
- parâmetros originais;
- postcheck;
- rollback validator;
- `rollback_preconditions` quando a reversão tiver pré-condições diferentes da ida;
- auditoria.

Nunca reutilize automaticamente `preconditions` da execução para o rollback: o estado pós-ação pode, por definição, invalidar a condição original.

Não adicione rollback “best effort” sem prova de estado.

## Seletores

Use SelectorKind para parâmetros que podem ser inventariados.

Não faça módulo chamar `input()`. A UI resolve seleção.

## Configuração corporativa

Nova allowlist deve:

- ter schema/validador de bootstrap;
- desabilitar apenas entrada inválida;
- nunca aceitar segredo no repositório;
- ter teste de entrada malformada.

## Banco

Mudança de schema:

1. incrementar SCHEMA_VERSION;
2. criar migration sequencial;
3. manter upgrade de bancos anteriores;
4. atualizar retention;
5. adicionar teste.

Schema atual: 4.

## Logging/auditoria

Persistência operacional deve usar redaction.

Não persista comando bruto ou segredo apenas por conveniência.

## Testes

```powershell
python -W error::SyntaxWarning -m compileall -q .
python -m ruff check .
python -m pytest -q
```

O CI executa mypy e coverage nos contratos endurecidos.

## Definition of Done

```text
contrato tipado
+ policy/preconditions
+ execução segura
+ postcheck
+ UNKNOWN quando evidência é insuficiente
+ rollback apenas quando real
+ auditoria redigida
+ teste de happy path
+ teste de falha
+ CI verde
+ documentação
```
