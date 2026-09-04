# Documentação técnica — Central N2 Workstation v5

Este diretório contém a documentação operacional e de engenharia da Central N2.

## Índice

- architecture.md — arquitetura v5, transportes, jobs, sessão lógica, persistência e apresentação.
- modules.md — catálogo funcional dos módulos.
- configuration.md — settings.json, settings.local.json, PsExec, GLPI, runtime e baseline.
- operations.md — manual de operação do suporte N2.
- security.md — privilégios, transportes, segredos, auditoria e ações críticas.
- troubleshooting.md — diagnóstico da própria Central e dos transportes.
- development.md — evolução do código, testes e CI.
- v5_complete.md — visão consolidada da geração v5.
- workstation_v2.md — documento histórico.

## Hierarquia de referência

Em caso de divergência:

1. código do master;
2. testes automatizados;
3. documentação v5;
4. documentos históricos.

A documentação não deve prometer comportamento que não exista no código.

## Público-alvo

- suporte N2;
- administradores Windows;
- desenvolvedores;
- infraestrutura;
- segurança/auditoria.

## Política de atualização

Mudanças de menu, transporte, configuração, segurança, output, timeout, módulo ou remediação devem atualizar a documentação no mesmo PR.
