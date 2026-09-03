# Central de Manutenção N2 — Workstation v2

> [!WARNING]
> **Documento histórico.** Esta página descreve a geração v2 e não é a referência operacional atual. Para a v3, comece por [`README.md`](README.md), [`operations.md`](operations.md) e [`architecture.md`](architecture.md).

## Escopo

A v2 é dedicada exclusivamente à administração, diagnóstico, manutenção e conformidade de estações de trabalho Windows. Operações de switch, NAC, OUI, MAC ACL e configuração Intelbras foram removidas do projeto.

## Princípio operacional

A ferramenta deve responder primeiro **o que está errado** e só depois oferecer a ação administrativa correspondente.

Fluxo esperado:

```text
Selecionar estação
      ↓
Pré-flight
      ↓
Snapshot de saúde
      ↓
Score + desvios
      ↓
Diagnóstico detalhado
      ↓
Remediação confirmada pelo operador
```

## Menu principal

1. Selecionar / alterar computador
2. Visão geral / Saúde
3. Diagnóstico / Hardware
4. Usuários e perfis
5. Rede da estação
6. Software
7. Processos e serviços
8. Windows Update
9. Segurança
10. Domínio / GPO
11. Impressoras
12. GLPI Agent
13. Disco / limpeza
14. Eventos
15. Energia / mensagens
16. Coletar diagnóstico
17. Compliance
18. Ações em lote

## Saúde da estação

O módulo `modules/health.py` coleta:

- CPU;
- RAM;
- espaço livre no disco C:;
- uptime;
- reinicialização pendente;
- serviços automáticos parados;
- Defender;
- Firewall;
- GLPI Agent.

O score começa em 100 e perde pontos conforme a gravidade dos desvios. O objetivo não é substituir monitoramento corporativo, mas oferecer ao técnico um indicador rápido de prioridade.

## Compliance

O baseline fica em `config/settings.json`:

```json
"compliance": {
  "min_disk_free_percent": 15,
  "max_uptime_days": 30,
  "defender_required": true,
  "firewall_required": true,
  "glpi_required": true,
  "pending_reboot_not_allowed": true
}
```

Ajuste esses valores à política real da organização. O score de compliance não deve ser tratado como certificação de segurança; ele representa somente os controles declarados no baseline.

## Segurança

O módulo consulta, quando disponível:

- Microsoft Defender;
- proteção em tempo real;
- atualização de assinatura;
- idade das varreduras;
- ameaças recentes;
- Firewall por perfil;
- BitLocker;
- TPM;
- Secure Boot;
- RDP;
- SMBv1;
- UAC.

A v2 não fornece comandos para desativar Defender ou Firewall.

## Windows Update

Recursos:

- histórico de hotfixes;
- busca de atualizações pendentes via Windows Update Agent;
- disparo de nova busca;
- reset controlado dos componentes BITS, Windows Update e Cryptographic Services.

O reset renomeia `SoftwareDistribution` e `catroot2` com timestamp em vez de apagá-los imediatamente, facilitando rollback/manual recovery.

## Usuários e perfis

Recursos:

- sessões abertas;
- administradores locais;
- perfis Win32_UserProfile;
- último uso;
- tamanho aproximado dos perfis;
- remoção de perfil por SID com confirmação.

A remoção de perfil é destrutiva e não deve ser usada sem validação do usuário conectado e de dados locais importantes.

## Impressoras

Recursos:

- impressoras instaladas;
- driver, porta e estado;
- filas de impressão;
- reinício do Spooler;
- limpeza das filas com confirmação.

## Domínio / GPO

Recursos:

- domínio atual;
- participação no domínio;
- secure channel;
- identificação do DC via `nltest`;
- status de hora via `w32tm`;
- `gpresult`;
- `gpupdate /force`;
- tentativa controlada de reparo do secure channel.

O reparo pode exigir permissões de domínio e deve ser testado com a política da organização.

## Disco e limpeza

A ferramenta pode:

- consultar capacidade e espaço livre;
- calcular tamanho aproximado dos perfis;
- estimar dados temporários recuperáveis;
- limpar TEMP do usuário/sistema e Lixeira.

A v2 **não apaga Downloads automaticamente** e não remove perfis como parte da limpeza genérica.

## Pacote de diagnóstico

`Coletar diagnóstico` gera um JSON local em:

```text
central_n2/reports/diagnostics/
```

A coleta inclui informações de computador/SO, discos, rede, serviços automáticos parados, eventos críticos, impressoras e processos com maior uso de memória.

A pasta `reports/` é ignorada pelo Git para evitar versionar dados reais das estações.

## Ações em lote

A v2 mantém:

- ping;
- GPUpdate;
- flush DNS;
- inventário GLPI;
- snapshot de saúde.

A concorrência permanece limitada para reduzir carga administrativa.

## Operações que exigem confirmação

- remoção de perfil;
- finalização de processo;
- parada/reinício de serviço;
- reset Winsock/TCP-IP;
- desabilitar Wi-Fi;
- remoção de software;
- reset de componentes do Windows Update;
- reparo de secure channel;
- limpeza de filas;
- limpeza de disco;
- reinício/desligamento.

## Testes

```powershell
cd central_n2
python -m pip install pytest
python -m pytest -q
```

Os testes da v2 validam o cálculo de saúde e o motor de compliance sem depender de uma estação remota.

## Limitações

- vários cmdlets dependem da versão/edição do Windows e dos módulos disponíveis;
- BitLocker, Defender e Secure Boot podem retornar `null` quando o recurso/cmdlet não estiver disponível;
- cálculo do tamanho de perfis pode ser demorado em perfis grandes;
- Windows Update Agent pode demorar dependendo do estado do serviço e da conectividade;
- WinRM continua sendo o transporte preferencial, com PsExec como fallback;
- a ferramenta pressupõe uso por equipe administrativa autorizada.
