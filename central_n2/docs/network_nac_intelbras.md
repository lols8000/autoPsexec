# Network / NAC Intelbras

## Objetivo

Este módulo reduz a presença de equipamentos não homologados conectados fisicamente à rede. Ele trabalha em duas fases separadas:

1. **Auditoria**: coleta a tabela MAC, normaliza os endereços, identifica OUI/fabricante e aponta equipamentos fora da allowlist.
2. **Planejamento**: gera uma proposta de MAC ACL para portas de acesso compatíveis, sem aplicar nem salvar automaticamente a configuração.

A separação é proposital. O operador deve observar a rede real antes de ativar qualquer política `deny-by-default`.

## Escopo inicial

Fabricantes esperados no ambiente:

- HP
- Positivo
- Daten
- Epson
- Ubiquiti / UniFi

Os OUIs **não são preenchidos automaticamente**. Eles devem ser cadastrados em `config/oui_allowlist.json` a partir do inventário real e de fontes confiáveis. Isso evita autorizar prefixos incorretos ou desatualizados.

## Arquivos

- `modules/network_nac.py`: parser, classificação OUI, perfis Intelbras, planejador ACL e cliente SSH.
- `network_nac_cli.py`: interface de auditoria e geração de plano.
- `config/oui_allowlist.json`: allowlist operacional.
- `config/oui_allowlist.example.json`: exemplo de formato.
- `tests/test_network_nac.py`: testes offline.
- `reports/network_nac/`: relatórios locais; não deve ser versionado.

## Execução

A partir da pasta `central_n2`:

```powershell
python .\network_nac_cli.py
```

## Auditoria por arquivo

No switch, execute manualmente o comando de exibição da tabela MAC correspondente ao modelo e salve a saída em um arquivo de texto. No módulo:

1. selecione `Auditar tabela MAC a partir de arquivo`;
2. informe o arquivo;
3. revise a classificação;
4. valide todos os itens marcados como `BLOQUEAR`.

O parser aceita formatos comuns de MAC:

- `00:11:22:33:44:55`
- `00-11-22-33-44-55`
- `0011.2233.4455`
- `001122334455`

## Auditoria via SSH

O cliente usa o OpenSSH do Windows em modo não interativo. A recomendação é autenticação por chave SSH.

A aplicação não armazena senha de switch.

Pré-requisitos:

- cliente OpenSSH disponível no Windows;
- SSH habilitado no switch;
- usuário administrativo específico para automação/auditoria;
- chave SSH protegida pelo usuário/Windows;
- conectividade somente a partir da rede administrativa.

## Perfis de switch

### Intelbras Série 3000

A documentação da Série 3000 descreve MAC ACL padrão com wildcard, IDs de ACL MAC padrão no intervalo 2001–3000 e aplicação por `mac access-group` em interface.

Esse é o único perfil em que a v1 habilita geração automática de allowlist por OUI `/24`.

Exemplo conceitual para OUI `00:11:22`:

```text
0011.2200.0000 0000.00FF.FFFF
```

Os primeiros 24 bits ficam fixos e os últimos 24 bits são ignorados pelo wildcard.

### Intelbras S2050G-A

A documentação confirma MAC ACL, regras e aplicação com `mac-access-group`. A v1 permite componentes para ACL exata, porém **não gera política OUI por wildcard automaticamente** neste perfil.

### Intelbras S-Series

A documentação confirma MAC ACL estendida e aplicação em interface. A v1 não usa esse perfil para geração OUI automática. Também deve ser respeitado o limite de regras da família/firmware.

## Regra de segurança principal

Nunca aplique a política de endpoint indiscriminadamente em:

- uplink para outro switch;
- trunk 802.1Q;
- porta de hypervisor;
- porta de Access Point que transporte MACs de clientes;
- telefone IP com computador encadeado;
- bridge;
- firewall/roteador;
- porta de servidor que represente múltiplos MACs.

Essas portas devem ser classificadas como infraestrutura e tratadas por política própria.

## Fluxo recomendado de implantação

### Fase 1 — somente auditoria

Execute durante alguns dias/turnos e consolide:

- MAC;
- OUI;
- fabricante;
- switch;
- porta;
- VLAN;
- usuário/local físico;
- justificativa do equipamento.

Não bloqueie nada nesta fase.

### Fase 2 — saneamento

Classifique cada ocorrência como:

- homologado;
- exceção documentada;
- infraestrutura;
- equipamento indevido;
- desconhecido e pendente de análise.

### Fase 3 — bancada

Escolha uma única porta de acesso que não seja crítica.

1. faça backup da configuração do switch;
2. garanta acesso por console ou caminho alternativo;
3. conecte um equipamento homologado;
4. gere o plano;
5. revise cada comando;
6. aplique manualmente;
7. valide DHCP, DNS, gateway e aplicações;
8. conecte um equipamento não homologado;
9. confirme que ele é bloqueado;
10. remova a ACL se houver comportamento inesperado.

### Fase 4 — piloto

Aplique em um pequeno grupo de portas de usuário.

### Fase 5 — expansão

Somente depois do piloto bem-sucedido, expanda por setores.

## DHCP não substitui a ACL

Negar lease DHCP ajuda na governança, mas não impede um dispositivo de configurar IP estático. A política de switch atua em camada 2 e reduz esse espaço.

A arquitetura recomendada é:

```text
porta física
    ↓
MAC ACL / Port Security
    ↓
VLAN
    ↓
DHCP / DHCP Snooping
    ↓
Firewall
    ↓
recursos autorizados
```

## Limitações de OUI

OUI não é autenticação forte.

Um usuário tecnicamente preparado pode alterar/spoofar o endereço MAC. Além disso:

- alguns dispositivos usam MAC localmente administrado/randomizado;
- um fabricante pode possuir dezenas de OUIs;
- placas de rede substituídas podem ser de outro fabricante;
- docks USB/Ethernet podem apresentar OUI diferente do computador.

Para autenticação forte de endpoint, a evolução recomendada é 802.1X/RADIUS/NAC baseado em identidade/certificado.

## MAC localmente administrado

O módulo identifica o bit `locally administered` do MAC. Esses endereços são marcados como não homologados quando não estiverem na lista de exceções exatas.

## Geração de ACL

A opção de geração cria apenas um **plano de configuração em arquivo**. A v1 não salva a configuração no switch automaticamente.

Isso é deliberado para impedir:

- bloqueio acidental de uplink;
- perda de acesso administrativo;
- persistência de uma ACL incorreta após reboot;
- implantação de uma sintaxe incompatível com firmware/modelo.

## Rollback

Antes de qualquer aplicação real, documente o rollback exato para o modelo/firmware. Em geral, o plano deve incluir:

- remoção da associação da ACL das interfaces afetadas;
- remoção/edição da ACL;
- restauração do backup de configuração quando necessário.

Não dependa apenas de SSH para rollback. Em mudanças de controle de acesso, console/OOB é a opção mais segura.

## Testes automatizados

Os testes não acessam switches reais. Eles validam:

- normalização de MAC;
- normalização de OUI;
- conversão para formato Intelbras;
- wildcard OUI da Série 3000;
- detecção de MAC localmente administrado;
- parsing de tabelas MAC;
- deduplicação;
- classificação autorizado/não autorizado;
- agrupamento por porta;
- recusa de ACL vazia;
- faixa de ID da ACL da Série 3000;
- recusa de política OUI em perfil não validado.

Execute:

```powershell
python -m pytest -q
```

## Critério para considerar o módulo pronto para produção

O módulo só deve deixar o estado de laboratório quando houver:

- modelo e firmware dos switches inventariados;
- OUIs reais validados;
- portas de infraestrutura catalogadas;
- backup testado;
- rollback testado;
- bancada concluída;
- piloto concluído;
- procedimento de exceção documentado;
- acesso administrativo alternativo disponível.
