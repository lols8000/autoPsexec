# Central de Manutenção N2

Nova geração do projeto `autoPsexec`, preservando os scripts históricos existentes na raiz do repositório.

## Objetivo

Transformar o antigo utilitário baseado em PsExec em uma central modular para suporte e administração remota de estações Windows, com melhor usabilidade, auditoria, diagnóstico e possibilidade de execução em lote.

## Arquitetura planejada

```text
central_n2/
├── core/       # execução, resultados, logging e utilitários centrais
├── remote/     # WinRM, PowerShell Remoting e fallback PsExec
├── modules/    # diagnóstico, rede, software, serviços, GLPI e manutenção
├── ui/         # interface da aplicação
├── config/     # configurações e catálogos
└── tests/      # testes automatizados
```

## Diretrizes

- Preservar o código legado para referência histórica.
- Preferir dados estruturados a parsing frágil de texto.
- Evitar `shell=True` quando não for necessário.
- Usar PowerShell/CIM para recursos Windows modernos.
- Manter PsExec como mecanismo de fallback quando útil.
- Registrar ações administrativas em log.
- Exigir confirmação para operações destrutivas ou disruptivas.
- Preparar execução em um único host e em lote.

## Próximas etapas

1. Criar o `RemoteExecutor`.
2. Implementar seleção e diagnóstico inicial do host.
3. Migrar as funções úteis do script antigo.
4. Substituir chamadas WMIC por PowerShell/CIM.
5. Adicionar logging e tratamento padronizado de erros.
6. Criar interface de console mais intuitiva.
7. Adicionar testes.
