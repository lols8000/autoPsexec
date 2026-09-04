# Segurança — Central N2 Workstation v5

## Modelo de confiança

A Central é ferramenta administrativa privilegiada. Pressupõe operador autorizado, estação administrativa confiável, rede permitida e hosts autorizados.

## UAC

main.py solicita elevação quando necessário.

## Segredos

Não persistir em código público:

- senha;
- token;
- API key;
- Authorization;
- chaves privadas;
- credenciais de domínio.

Use settings.local.json para configuração privada.

## Repositório público

Não commitar:

- IPs internos desnecessários;
- inventários reais;
- nomes de usuários;
- logs;
- relatórios;
- tokens;
- paths sensíveis.

## WinRM

Use conforme política corporativa. Evite TrustedHosts=* e desabilitação de controles apenas para “fazer funcionar”.

## PsExec

PsExec é fallback legítimo, mas deve ser homologado.

Riscos/considerações:

- depende de SMB/ADMIN$;
- pode ser bloqueado por EDR;
- pode executar em contexto SYSTEM;
- programas interativos podem mudar de comportamento;
- Winget pode não estar disponível.

A Central não baixa PsExec automaticamente.

## Ruído de transporte

A limpeza de Starting..., exit code 0 e CLIXML só ocorre em execução bem-sucedida. Em falha, mensagens são preservadas para diagnóstico.

## Ações destrutivas/disruptivas

Exemplos:

- remoção de perfil;
- kill de processo;
- stop de serviço;
- reset de rede;
- limpeza de fila;
- reboot/shutdown;
- Component Store cleanup;
- secure channel repair.

Devem mostrar host, exigir confirmação quando aplicável, possuir timeout, retornar resultado e ser auditáveis.

## Concorrência

JobManager serializa mutações por host para evitar DISM + cleanup + reboot simultâneos.

## Timeout

Timeout não é cancelamento remoto garantido.

## Auditoria

AuditLogger registra correlation_id e sanitiza campos sensíveis antes de persistir.

## Persistência

SQLite, relatórios e logs contêm dados operacionais e devem ter ACL adequada.

## Sysinternals

Use pacote homologado. A Central não faz download silencioso.

## GLPI API

Tokens devem permanecer em settings.local.json. A API é desabilitada por padrão.

## Entrada do operador

Evitar terminal remoto livre e concatenação de comando arbitrário sem validação. Preferir parâmetros estruturados e shell=False.

## Incidente

Em suspeita de abuso:

1. interromper uso;
2. preservar logs;
3. identificar estação administrativa;
4. identificar alvos;
5. revisar correlation_id/jobs;
6. revisar credenciais;
7. acionar segurança.
