# Operação — Central N2 Workstation 5.2.0

## Regra principal

```text
EVIDÊNCIA → DIAGNÓSTICO → PLANO → EXECUÇÃO → VALIDAÇÃO → REGISTRO
```

A Central não deve ser usada como lançador indiscriminado de comandos.

## 1. Selecionar estação

Prefira hostname/FQDN.

A seleção:

1. cria novo AttendanceContext;
2. gera correlation_id;
3. executa preflight;
4. seleciona transporte;
5. coleta capabilities;
6. coleta snapshot inicial quando possível.

Não use a Central de Execuções se a sessão não estiver READY.

## 2. Abrir a Central de Execuções

Menu:

```text
[28] Central de Execuções
```

Também há atalhos `90+` nos menus de diagnóstico.

A tela mostra categorias e permite:

- escolher domínio;
- buscar ação por texto/tag;
- desfazer a última ação reversível do atendimento.

## 3. Escolher parâmetros

Parâmetros são tipados.

Quando existe seletor, a Central inventaria o host e mostra objetos reais. Perfis usam inventário leve, sem cálculo recursivo de tamanho, e sessões usam dados estruturados de sessões interativas em vez de parsing textual de `quser`. Exemplo:

```text
Serviço:
1 - Print Spooler | Spooler | Running | Automatic
2 - Windows Update | wuauserv | Running | Manual
M - Informar manualmente
```

Use entrada manual apenas quando o objeto não aparecer.

## 4. Ler o plano

Antes da confirmação:

```text
✓ [PASS] Transporte psexec permitido.
✓ [PASS] Contexto administrativo confirmado.
✓ [PASS] Capability PnPUtil disponível.
✗ [FAIL] Winget indisponível.

Plano: BLOQUEADO
```

Ação bloqueada não é executada.

## 5. Confirmar

Ações normais: `SIM`.

Ações de risco elevado:

```text
EXECUTAR <hostname>
```

Leia impacto, risco, transportes, retry e rollback antes de confirmar.

## 6. Resultado

Estados:

- PASS — objetivo comprovado;
- FAIL — falha comprovada;
- UNKNOWN — estado final não comprovado.

`CommandResult.indeterminate` nunca deve ser tratado como falha simples para repetir automaticamente.

## 7. Disconnect temporário

Exemplo: restart de NIC ou reboot.

```text
comando entregue
 ↓
conexão cai
 ↓
Central aguarda
 ↓
SessionManager refresh
 ↓
READY novamente
 ↓
postcheck
 ↓
PASS/FAIL/UNKNOWN
```

Se o host não voltar dentro do prazo, o resultado é UNKNOWN.

## 8. Shutdown

Shutdown usa disconnect terminal. A Central não espera a máquina voltar.

## 9. Rollback

Quando disponível, o menu 28 mostra:

```text
U - Desfazer última execução reversível
```

Confirmação:

```text
DESFAZER <hostname>
```

Rollback é validado e persistido como nova execução ligada à original. Um rollback bem-sucedido consome a disponibilidade de rollback do registro pai. Guardas do rollback são avaliadas separadamente das preconditions da execução original.

A Central mantém uma pilha LIFO das execuções reversíveis do atendimento. Se uma ação não reversível for executada depois, a reversível anterior continua disponível.

Não existe rollback automático para operações irreversíveis.

## 10. Retry

Padrão: `PRE_EXECUTION_ONLY`.

A Central pode repetir automaticamente somente quando a falha é comprovadamente anterior à entrega da ação. O número máximo e o atraso entre tentativas são definidos no contrato da ação.

`indeterminate` nunca é repetido automaticamente.

Depois de timeout/perda de sessão com entrega incerta:

1. consulte o estado;
2. verifique se a ação pode ter sido entregue;
3. use o postcheck pertinente;
4. só repita quando houver evidência suficiente.

## 11. Catálogos corporativos

No bootstrap, pacotes/certificados/Registro são validados.

Entrada inválida é desabilitada e aparece como warning.

Não corrija isso liberando shell ou instalador arbitrário.

## 12. Histórico

Menu 21 mostra snapshots e execuções recentes.

Execuções registram:

- action;
- validation state;
- operador;
- transporte;
- risco;
- duração;
- correlation_id;
- rollback linkage.

## 13. GLPI / relatório

O relatório deve refletir o resultado validado, não apenas o comando executado.

Ao registrar no GLPI, use correlation_id para ligar diagnóstico, jobs e execuções.
