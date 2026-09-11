# Documentação técnica — Central N2 Workstation 5.2.0

## Índice

- [architecture.md](architecture.md) — arquitetura, ExecutionEngine, policy, transportes, recovery e persistência.
- [configuration.md](configuration.md) — configuração efetiva, allowlists e validação no bootstrap.
- [modules.md](modules.md) — módulos funcionais e catálogos de execução.
- [operations.md](operations.md) — runbook operacional do suporte N2.
- [security.md](security.md) — confiança, privilege, retry, rollback e auditoria.
- [troubleshooting.md](troubleshooting.md) — falhas de transporte, policy e recovery.
- [development.md](development.md) — como adicionar ações e critérios de DoD.
- [release_and_distribution.md](release_and_distribution.md) — CI, build, assinatura, hashes e release.
- [v5_complete.md](v5_complete.md) — visão consolidada da geração 5.2.
- [workstation_v2.md](workstation_v2.md) — referência histórica.

## Fonte de verdade

Em caso de divergência:

1. código da release em execução;
2. testes automatizados;
3. documentação da mesma versão;
4. documentos históricos.

## Contratos críticos da 5.2

- `AttendanceContext` é a fonte única do atendimento;
- `READY_WINRM` exige execução autenticada;
- mutação com entrega incerta é `indeterminate` e não sofre retry cego;
- `ExecutionAction` descreve risco, transportes, privilege, capabilities, retry e rollback;
- `ExecutionPolicy` bloqueia incompatibilidades antes da confirmação;
- disconnect temporário exige recovery antes do postcheck;
- rollback só existe quando o estado anterior pode ser restaurado com evidência;
- allowlists corporativas são validadas no bootstrap;
- SQLite schema 4 mantém auditoria rica e redigida;
- CI verde é requisito de promoção;
- segredos nunca entram no repositório.
