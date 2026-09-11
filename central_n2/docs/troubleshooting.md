# Troubleshooting — Central N2 Workstation 5.2.0

## 1. Host resolve, ping responde, WinRM falha

Teste:

```powershell
Test-NetConnection HOST -Port 5985
Test-NetConnection HOST -Port 445
dir \\HOST\admin$
```

Se 5985 falhar e 445/ADMIN$ funcionarem, valide PsExec.

A Central pode operar em READY_PSEXEC sem alterar política de WinRM.

## 2. WinRM por IP falha

Em domínio, prefira hostname/FQDN.

Não use `TrustedHosts=*` como solução genérica.

## 3. PsExec não aparece

Valide:

```powershell
Test-Path C:\Sysinternals\PsExec.exe
C:\Sysinternals\PsExec.exe -accepteula \\HOST hostname
```

ADMIN$ disponível não garante acesso ao Service Control Manager.

## 4. Ação aparece como BLOQUEADA

Leia os PolicyChecks.

Exemplos:

```text
✗ Transporte psexec não permitido.
✗ Capability Winget indisponível.
✗ Contexto USER_CONTEXT não confirmado.
✗ Precondition: destino já existe.
```

Não contorne o bloqueio editando o código em produção. Corrija o requisito ou escolha ação compatível.

## 5. Winget bloqueado em PsExec

É esperado em alguns contextos.

Winget pode depender do perfil do usuário e não estar disponível sob SYSTEM/serviço remoto.

Use transporte/contexto permitido pela policy.

## 6. Resultado UNKNOWN

UNKNOWN significa falta de evidência final.

Causas comuns:

- timeout;
- host caiu e não voltou;
- postcheck falhou;
- comando terminou sem resultado estruturado;
- recurso mudou para estado não esperado.

Não trate UNKNOWN como PASS ou FAIL automático.

## 7. Resultado indeterminate

A ação pode ter sido entregue, mas a confirmação foi perdida.

Procedimento:

1. não repetir imediatamente;
2. consultar o estado;
3. usar diagnóstico/postcheck pertinente;
4. repetir apenas com evidência suficiente.

## 8. Recovery falhou

Para DisconnectMode.TEMPORARY:

```text
RECUPERAÇÃO: NÃO RECUPERADO
```

Verifique:

- host voltou a responder;
- DNS;
- 445/5985;
- ADMIN$;
- PsExec;
- boot em andamento;
- DHCP/endereço alterado.

O resultado deve permanecer UNKNOWN se a Central não comprovar o estado.

## 9. Rollback não aparece

Rollback só existe para ações reversíveis.

Também deixa de estar disponível após rollback validado com PASS.

Operações irreversíveis não oferecem opção U.

## 10. Catálogo corporativo não aparece

No startup, procure warnings:

```text
packages.chave: ...
certificates.chave: ...
registry_actions.chave: ...
```

Entrada inválida é removida do catálogo efetivo.

## 11. Banco falha no startup

SQLite executa quick_check e migrations sequenciais.

Não altere `PRAGMA user_version` manualmente.

Faça backup do arquivo antes de intervenção.

## 12. Driver inventory parece bruto

A saída atual usa JSON estruturado e a UI deve renderizar campos de maneira amigável. Se aparecer CLIXML/ruído de transporte, registre o caso com correlation_id e transporte utilizado.

## 13. Como coletar evidência

Use:

- menu 19 para conectividade/capabilities;
- menu 21 para histórico;
- menu 23 para jobs;
- relatório de suporte;
- correlation_id exibido na UI.

Não publique logs reais no repositório público.
