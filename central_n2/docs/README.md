# Documentação técnica — Central N2 Workstation v3

Este diretório contém a documentação operacional e de engenharia da Central N2.

A documentação foi separada por responsabilidade para evitar um único arquivo extenso, difícil de manter e sujeito a divergências entre código, operação e configuração.

## Índice

### [`architecture.md`](architecture.md)

Arquitetura da aplicação, camadas, fluxo de execução, transportes remotos, `CommandResult`, runner responsivo, threads, timeouts, logging e critérios para expansão do projeto.

### [`modules.md`](modules.md)

Catálogo dos módulos disponíveis, finalidade, tipo de operação, principais comandos utilizados e cuidados de execução.

### [`configuration.md`](configuration.md)

Descrição detalhada do `config/settings.json`, configuração de PsExec, Sysinternals, GLPI, software, compliance e parâmetros da interface.

### [`operations.md`](operations.md)

Manual do operador N2: inicialização, seleção de estação, interpretação de status, fluxo de diagnóstico, uso dos menus, remediação e geração de evidência.

### [`security.md`](security.md)

Modelo de privilégios, superfície de risco, confirmação de ações, tratamento de credenciais, logging, uso de Sysinternals e cuidados com repositório público.

### [`troubleshooting.md`](troubleshooting.md)

Diagnóstico de problemas da própria Central: Python, UAC, DNS, WinRM, PsExec, `ADMIN$`, timeout, Winget, GLPI, Sysinternals e comandos remotos.

### [`development.md`](development.md)

Guia para evolução do código: convenções, novos módulos, contratos de retorno, threads, timeouts, testes, CI e checklist de PR.

### [`workstation_v2.md`](workstation_v2.md)

Documento histórico da geração v2. Deve ser tratado como referência de evolução; a documentação normativa atual é a v3 descrita pelos documentos acima.

---

## Hierarquia de referência

Em caso de divergência, a ordem de autoridade é:

1. comportamento efetivamente implementado no código do `master`;
2. testes automatizados;
3. documentação v3 neste diretório;
4. documentos históricos de versões anteriores.

A documentação não deve prometer funcionalidades que ainda não existam no código.

---

## Público-alvo

- suporte N2;
- administradores Windows;
- desenvolvedores responsáveis pela Central;
- equipe de infraestrutura que precisa auditar o comportamento da ferramenta;
- responsáveis por segurança que precisam entender privilégios, transportes e ações executadas.

---

## Política de atualização

Toda mudança relevante deve atualizar a documentação correspondente no mesmo PR quando alterar:

- configuração;
- fluxo operacional;
- menu;
- comandos remotos;
- segurança;
- comportamento de timeout;
- dependências;
- requisitos de ambiente;
- novos módulos;
- ações destrutivas ou disruptivas.
