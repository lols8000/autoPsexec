# Documentação técnica — Central N2 Workstation 5.1.0

## Índice

- [architecture.md](architecture.md) — componentes, estado, transporte, fallback, jobs e persistência.
- [configuration.md](configuration.md) — configuração efetiva e precedência.
- [modules.md](modules.md) — catálogo funcional.
- [operations.md](operations.md) — runbook do suporte N2.
- [security.md](security.md) — modelo de confiança e controles.
- [troubleshooting.md](troubleshooting.md) — falhas da Central e transportes.
- [development.md](development.md) — regras de evolução, testes e DoD.
- [release_and_distribution.md](release_and_distribution.md) — CI, build, assinatura, hashes e release.
- [v5_complete.md](v5_complete.md) — visão consolidada da geração 5.1.
- [workstation_v2.md](workstation_v2.md) — referência histórica.

## Fonte de verdade

Em caso de divergência:

1. código do branch/release em execução;
2. testes automatizados;
3. documentação correspondente à mesma versão;
4. documentos históricos.

A documentação não deve prometer uma capacidade inexistente nem esconder uma limitação real.

## Contratos críticos da 5.1

- \`AttendanceContext\` é a fonte única de estado do atendimento;
- \`READY_WINRM\` exige \`Invoke-Command\` validado;
- leitura pode fazer fallback; mutação não pode ser repetida quando a entrega é incerta;
- resultado indeterminado exige validação antes de repetir;
- compliance diferencia PASS, FAIL, UNKNOWN e N/A;
- schema SQLite é versionado por migration;
- CI precisa ficar verde antes de promoção;
- segredos nunca entram no repositório.

## Público

- suporte N2;
- administradores Windows;
- infraestrutura;
- segurança/auditoria;
- mantenedores da Central.

## Política de atualização

Mudança de transporte, fallback, menu, configuração, avaliação, remediação, schema, build ou release deve atualizar a documentação no mesmo PR.
