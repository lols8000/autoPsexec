# Operação — Central N2 Workstation 5.1.0

## Regra principal

\`\`\`text
EVIDÊNCIA → DIAGNÓSTICO → REMEDIAÇÃO → VALIDAÇÃO → REGISTRO
\`\`\`

Não use a Central como lançador indiscriminado de comandos de reparo.

## 1. Iniciar

\`\`\`powershell
python .\central_n2\main.py
\`\`\`

Confirme a elevação UAC.

## 2. Selecionar estação

Prefira hostname/FQDN em domínio. A seleção cria um novo \`AttendanceContext\`, gera novo \`correlation_id\`, executa preflight e coleta snapshot inicial quando há transporte válido.

Leia o estado:

- \`READY_LOCAL\`: execução direta;
- \`READY_WINRM\`: WinRM autenticado;
- \`READY_PSEXEC\`: PsExec validado;
- qualquer outro: investigar antes de tentar remediação.

## 3. Cenário WinRM bloqueado

Exemplo:

\`\`\`text
DNS ............. OK
Ping ............ OK
TCP 5985 ........ FAIL
TCP 445 ......... OK
ADMIN$ .......... OK
PsExec .......... OK
Estado .......... READY_PSEXEC
\`\`\`

Isso é suportado. Não é necessário abrir 5985 apenas para atender pela Central se a política do ambiente permite PsExec.

## 4. WinRM por IP

WinRM pode falhar por autenticação ao usar IP. Em domínio, prefira hostname/FQDN.

Não use \`TrustedHosts=*\` como correção genérica.

## 5. Compliance

Interpretação:

| Estado | Ação |
| --- | --- |
| PASS | atende ao baseline |
| FAIL | desvio confirmado |
| UNKNOWN | não há evidência suficiente; investigar |
| N/A | controle não exigido pelo baseline |

Nunca trate UNKNOWN como prova de não conformidade.

## 6. Operações longas

SFC, DISM, inventário pesado e outras rotinas podem levar minutos. O JobManager mantém feedback e timeout.

**Timeout não é cancelamento remoto garantido.**

Antes de repetir uma operação pesada após timeout:

1. consulte o estado atual;
2. verifique processo/serviço/log pertinente;
3. confirme se a ação ainda está em andamento;
4. só então decida repetir.

## 7. Resultado indeterminado

Se uma ação mutável perder comunicação depois de possivelmente ter sido entregue, a Central marca o resultado como indeterminado.

Procedimento:

\`\`\`text
não repetir imediatamente
  ↓
consultar estado final do recurso
  ↓
comparar com objetivo da ação
  ↓
se necessário, executar nova remediação consciente
\`\`\`

O fallback automático para PsExec fica bloqueado nesse caso para evitar dupla execução.

## 8. Fluxos práticos

### Lentidão

\`\`\`text
Saúde → Performance → Disco/Startup → Playbook Lentidão → causa → remediação
\`\`\`

### Rede

\`\`\`text
interface → IP → gateway → DNS → TCP específico
\`\`\`

Renovação DHCP é disruptiva; use somente quando a causa justificar.

### Impressão

\`\`\`text
inventário → fila → Spooler → driver/porta/rede
\`\`\`

### Domínio/GPO

\`\`\`text
domínio/DC → horário → secure channel → gpresult → gpupdate/repair se necessário
\`\`\`

### Windows Update

\`\`\`text
status/pendências → causa → reset apenas se justificado → validação dos serviços
\`\`\`

### Disco cheio

\`\`\`text
uso → perfis → estimativa → limpeza segura
\`\`\`

A limpeza guiada atua em temporários e **não toca a Lixeira nem Downloads**.

## 9. Remediações guiadas

Antes de confirmar, leia impacto e observação de rollback.

Validações:

- Spooler: serviço deve estar Running;
- limpeza: retorno estruturado deve confirmar recuperação;
- Windows Update: serviços originalmente ativos devem ser restaurados;
- GPUpdate: execução deve concluir e GPResult pós-ação deve ser coletado.

Validação UNKNOWN significa que o estado final não foi suficientemente comprovado.

## 10. Histórico e diff

Use o menu de histórico para comparar os dois snapshots de saúde mais recentes. O SQLite persiste dados com \`correlation_id\`.

## 11. Relatório

Gere após o atendimento. O relatório registra:

- host;
- usuário quando disponível;
- problema;
- diagnóstico;
- ações;
- validação;
- resultado;
- correlation_id.

## 12. GLPI

A API é opcional. O envio exige relatório do atendimento atual e o contexto deve pertencer ao host selecionado.

## 13. Jobs

O menu Jobs mostra estado, classe, duração e correlation_id. Mutações são serializadas por host.

## 14. Checklist de encerramento

\`\`\`text
[ ] causa ou hipótese registrada
[ ] ação executada apenas quando justificada
[ ] resultado não ficou indeterminado sem investigação
[ ] validação pós-ação realizada
[ ] relatório gerado
[ ] chamado atualizado quando aplicável
\`\`\`
