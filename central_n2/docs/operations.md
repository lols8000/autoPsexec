# Manual operacional — Central N2 Workstation v3

## 1. Objetivo do manual

Este documento descreve como o suporte N2 deve operar a Central de forma previsível e segura.

A ferramenta deve ser utilizada com a seguinte sequência mental:

```text
Identificar host
    ↓
Confirmar conectividade
    ↓
Coletar saúde
    ↓
Classificar sintoma
    ↓
Diagnosticar
    ↓
Remediar
    ↓
Validar novamente
    ↓
Registrar evidência
```

Evite começar por uma ação de reparo apenas porque ela existe no menu.

---

## 2. Inicialização

Na raiz do repositório:

```powershell
python .\central_n2\main.py
```

A aplicação:

1. valida Windows;
2. valida configuração;
3. solicita elevação administrativa via UAC se necessário;
4. inicializa executor, logger e UI responsiva.

### Saída esperada

```text
╔════════════════════════════════════════════════════╗
║        CENTRAL N2 WORKSTATION — V3 RESPONSIVA     ║
╚════════════════════════════════════════════════════╝

Alvo: nenhum
```

---

## 3. Seleção da estação

Escolha:

```text
[1] Selecionar estação
```

Informe hostname ou IP.

Preferência operacional: use hostname quando DNS e inventário corporativo estiverem consistentes.

### Primeiro retorno

A UI executa snapshot de saúde com feedback visual:

```text
| Pré-flight da estação...  1.2s
```

Quando houver resultado:

```text
✓ SUCESSO [winrm]
```

ou fallback:

```text
✓ SUCESSO [psexec]
```

ou falha explícita.

---

## 4. Como interpretar o heartbeat

Durante uma operação longa:

```text
| SFC /scannow...  18.4s
```

Isso significa:

- a UI principal continua respondendo;
- o worker continua aguardando resultado;
- o tempo decorrido continua sendo atualizado;
- a Central não concluiu ainda a operação.

### Heartbeat não significa progresso percentual

O spinner indica **atividade da Central**, não porcentagem real do DISM/SFC/comando remoto.

Não interprete:

```text
30 segundos
```

como “30%”.

---

## 5. Status finais

### Sucesso

```text
✓ SUCESSO [winrm] — 8254 ms
```

O transporte entre colchetes ajuda a diagnosticar o caminho utilizado.

### Falha

```text
✗ FALHA [psexec]
Erro: Access denied
```

Leia `stderr` antes de repetir a ação.

### Timeout

```text
✗ TIMEOUT
Operação 'DISM RestoreHealth' excedeu 3600s
```

Um timeout local significa que a Central deixou de aguardar o resultado.

**Não significa necessariamente que o processo remoto foi encerrado.**

Antes de executar novamente uma operação pesada, valide se o processo ainda está ativo na estação.

---

## 6. Saúde / Compliance

Menu:

```text
[2] Saúde / Compliance
```

Use como triagem inicial.

Exemplo:

```text
Score: 72/100
CPU: 48%
RAM: 81%
Disco livre: 7%
Uptime: 44 dias

[HIGH] Disco C: abaixo do mínimo
[MEDIUM] Reinicialização pendente
```

### Interpretação

O score prioriza investigação, mas não determina sozinho a causa do chamado.

Exemplo: disco crítico pode ser relevante para lentidão, mas não explica necessariamente uma falha de impressão de rede.

---

## 7. Performance

Menu:

```text
[3] Performance
```

### Amostragem rápida

Use para triagem.

### Amostragem detalhada

Use quando a lentidão é intermitente ou não aparece em uma fotografia instantânea.

Procure:

- CPU média e pico;
- RAM;
- atividade de disco;
- tráfego de rede;
- processos dominantes.

### Evite

Rodar amostragem continuamente como monitoramento permanente. O módulo é uma ferramenta de diagnóstico pontual.

---

## 8. Reparo do Windows

Menu:

```text
[4] Reparo do Windows
```

### Sequência recomendada

Para suspeita de corrupção:

```text
DISM CheckHealth
      ↓
DISM ScanHealth
      ↓
DISM RestoreHealth, se necessário
      ↓
SFC /scannow
```

A ordem pode variar conforme diagnóstico, mas evite executar todas as opções automaticamente “por garantia”.

### Component Store cleanup

É remediação. Use quando houver justificativa operacional.

### CHKDSK `/scan`

É uma verificação online. Se indicar problema que exija correção offline/reboot, trate isso separadamente.

### WMI repository

O módulo atual verifica consistência; não faça reset/rebuild agressivo sem diagnóstico adicional.

---

## 9. Hardware / Drivers / Dispositivos

Menu:

```text
[5] Hardware / Drivers / Dispositivos
```

Use em casos de:

- placa de rede sem funcionar;
- USB não reconhecido;
- áudio/vídeo com erro;
- dispositivo Code 10/Code 43;
- driver suspeito;
- pós-formatação;
- necessidade de backup de drivers.

### Exportar drivers

A exportação grava arquivos na estação alvo. Confirme espaço e destino antes de uso em massa.

---

## 10. Inicialização / Tarefas

Menu:

```text
[6] Inicialização / Tarefas
```

Útil para:

- boot/login lento;
- programas abrindo sozinhos;
- software persistente;
- tarefas corporativas falhando;
- automações que deixaram de executar.

### Tarefa com falha

`LastTaskResult != 0` é um indicador, não diagnóstico completo. Consulte histórico/eventos da tarefa quando necessário.

---

## 11. Crashes / BSOD

Menu:

```text
[7] Crashes / BSOD
```

Use para:

- reinicialização inesperada;
- tela azul;
- aplicação fechando;
- crash recorrente.

A Central localiza evidências e dumps, mas análise aprofundada de dump pode exigir WinDbg/N3.

### Fluxo recomendado

```text
Identificar data/hora
      ↓
BugCheck / Application Error
      ↓
Localizar dump
      ↓
Coletar pacote
      ↓
Escalonar, se necessário
```

---

## 12. Segurança

Menu:

```text
[8] Segurança
```

Valide:

- Defender;
- assinatura;
- Firewall;
- BitLocker;
- TPM;
- Secure Boot;
- RDP;
- SMBv1;
- UAC;
- ameaças recentes.

A Central não deve ser utilizada para burlar controles corporativos.

---

## 13. Rede

Menu:

```text
[9] Rede
```

### Diagnóstico recomendado

```text
Interface
   ↓
IP
   ↓
Gateway
   ↓
DNS
   ↓
ARP/conexões
   ↓
teste específico
```

### Flush DNS

Baixo impacto relativo; use quando houver suspeita de cache incorreto.

### Renovar DHCP

Pode alterar IP e interromper conectividade temporariamente.

### Reset Winsock/TCP-IP

Ação mais agressiva. Use somente após diagnóstico e considere necessidade de reboot.

---

## 14. Usuários / Perfis

Menu:

```text
[10] Usuários / Perfis
```

Verifique:

- sessão ativa;
- administradores locais;
- tamanho de perfil;
- perfis antigos;
- SID correto antes de qualquer remoção.

### Remoção de perfil

É uma ação destrutiva.

Antes de remover:

```text
[ ] usuário não está logado
[ ] dados necessários foram preservados
[ ] SID foi validado
[ ] host correto foi validado
[ ] chamado autoriza a ação
```

---

## 15. Software / GLPI

Menu:

```text
[11] Software / GLPI
```

### Software

Use inventário antes de instalação/remoção.

### Winget

Se falhar remotamente, considere diferença entre contexto interativo do usuário e contexto SYSTEM/PsExec.

### GLPI

Antes de instalar/reparar, `installer_source` precisa estar configurado localmente.

---

## 16. Impressoras

Menu:

```text
[12] Impressoras
```

Fluxo recomendado:

```text
Inventário
   ↓
Fila
   ↓
Spooler
   ↓
conectividade/driver
   ↓
remediação
```

Limpar fila remove jobs pendentes. Confirme impacto com o usuário quando aplicável.

---

## 17. Domínio / GPO

Menu:

```text
[13] Domínio / GPO
```

Use para:

- GPO não aplicada;
- confiança quebrada;
- autenticação de domínio;
- horário incorreto;
- DC inesperado.

### Secure channel repair

É remediação. Exige contexto correto e deve ser executado somente se o canal realmente estiver quebrado.

---

## 18. Disco / Armazenamento / Bateria

Menu:

```text
[14] Disco / Armazenamento / Bateria
```

### Disco lógico

Observe espaço livre.

### Perfis

Use análise de tamanho antes de limpar.

### Limpeza

A rotina segura atual remove TEMP e lixeira; não apaga Downloads automaticamente.

### Disco físico

`HealthStatus` depende do que o stack de armazenamento e firmware expõem ao Windows.

### Bateria

Dados podem estar ausentes em desktops ou em equipamentos cujo firmware não fornece informações completas.

---

## 19. Ferramentas avançadas

Menu:

```text
[15] Ferramentas avançadas
```

Inclui:

- certificados;
- unidades mapeadas;
- shares locais;
- proxy;
- ativação Windows;
- logons recentes.

Essas consultas são úteis para problemas corporativos que não aparecem em “hardware/rede” de forma óbvia.

---

## 20. Sysinternals

Menu:

```text
[16] Sysinternals
```

Dependência: ferramentas precisam existir no diretório configurado.

### Autorunsc

Diagnóstico avançado de inicialização/persistência.

### ProcDump

Gera dump de um processo. Use quando houver crash/travamento que precisa de evidência para análise.

### Handle

Identifica qual processo mantém determinado arquivo/recurso aberto.

### Sigcheck

Consulta assinatura/hash de executável.

---

## 21. Pacote de diagnóstico

Menu:

```text
[17] Pacote de diagnóstico
```

Use antes de escalonar para N3 ou quando o chamado precisar de evidência técnica.

Os relatórios podem conter dados internos. Trate-os conforme política da organização.

---

## 22. Energia / Processos / Serviços

Menu:

```text
[18] Energia / Processos / Serviços
```

### Processos

Finalizar processo pode causar perda de dados não salvos.

### Serviços

Antes de parar/reiniciar um serviço, identifique dependências e impacto.

### Reiniciar/desligar

Confirmar:

```text
[ ] host correto
[ ] usuário foi avisado
[ ] janela de manutenção permite
[ ] não há tarefa crítica ativa
```

---

## 23. Quando interromper a Central

`Ctrl+C` encerra a espera local e a aplicação trata `KeyboardInterrupt`.

No entanto, uma operação remota já iniciada pode continuar executando.

Para ações longas, depois de interromper a UI, valide estado remoto antes de repetir o comando.

---

## 24. Fluxos práticos

### Lentidão

```text
Saúde
 ↓
Performance
 ↓
Disco / startup
 ↓
Eventos
 ↓
reparo somente se houver evidência
```

### Sem rede

```text
Interfaces
 ↓
IP/Gateway/DNS
 ↓
DHCP
 ↓
teste TCP/DNS
 ↓
flush/renew/reset somente conforme causa
```

### Impressão

```text
Inventário
 ↓
Fila
 ↓
Spooler
 ↓
rede/driver
```

### BSOD

```text
Crashes/BSOD
 ↓
BugCheck
 ↓
Dump
 ↓
Pacote de diagnóstico
```

### Problema de domínio

```text
Domínio/GPO
 ↓
DC
 ↓
horário
 ↓
secure channel
 ↓
gpresult
 ↓
repair/gpupdate se necessário
```

---

## 25. Princípio operacional

A ferramenta existe para reduzir tentativa e erro.

Prefira sempre:

```text
EVIDÊNCIA → DIAGNÓSTICO → REMEDIAÇÃO → VALIDAÇÃO
```

em vez de:

```text
SFC + DISM + RESET + REBOOT “para ver se resolve”
```
