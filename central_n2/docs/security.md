# Segurança — Central N2 Workstation v3

## 1. Modelo de confiança

A Central N2 é uma ferramenta administrativa. Por definição, ela opera com privilégios capazes de consultar e alterar estações Windows.

Ela deve ser tratada como software de administração privilegiada, não como utilitário comum de usuário.

O modelo pressupõe:

- operador autorizado;
- estação administrativa confiável;
- credenciais/contexto com privilégio compatível;
- rede corporativa confiável ou canal protegido;
- hosts alvo previamente autorizados.

---

## 2. Elevação administrativa

`central_n2/main.py` verifica privilégio administrativo e solicita elevação via UAC quando necessário.

A elevação não deve ser removida apenas para evitar prompts, pois várias funções dependem de contexto administrativo.

---

## 3. Credenciais

A Central não deve persistir:

- senhas;
- hashes de senha;
- tokens;
- API keys;
- chaves privadas;
- credenciais de domínio;
- segredo de proxy.

### Regra

Credenciais devem vir do contexto autorizado do operador/sessão ou de mecanismos corporativos apropriados, nunca de constantes no código.

---

## 4. Repositório público

O repositório é público.

Portanto, não devem ser commitados:

- IPs internos desnecessários;
- nomes de compartilhamentos privados;
- topologia real;
- inventário de estações;
- nomes de usuários reais;
- logs reais;
- pacotes de diagnóstico;
- chaves/tokens;
- caminhos internos sensíveis.

O `glpi.installer_source` público permanece vazio intencionalmente.

---

## 5. WinRM

WinRM/PowerShell Remoting é o transporte preferencial.

A configuração de WinRM deve obedecer à política do ambiente.

Evite configurações permissivas genéricas apenas para “fazer funcionar”, como ampliar TrustedHosts indiscriminadamente ou desabilitar controles de autenticação sem necessidade.

A configuração segura é responsabilidade da infraestrutura do ambiente.

---

## 6. PsExec

PsExec é fallback e possui características próprias:

- pode depender de SMB/RPC;
- pode executar em contexto SYSTEM;
- pode ser bloqueado por EDR/antivírus/política;
- comportamento de programas interativos pode mudar em contexto remoto;
- Winget pode não estar disponível da mesma forma em SYSTEM.

Use versão confiável e controle seu caminho local.

---

## 7. Ações destrutivas/disruptivas

Exemplos:

- remover perfil;
- finalizar processo;
- parar serviço;
- limpar fila de impressão;
- resetar rede;
- reiniciar;
- desligar;
- Component Store cleanup;
- secure channel repair;
- gerar dump em produção de usuário.

Essas ações devem:

1. mostrar host alvo;
2. solicitar confirmação quando apropriado;
3. possuir timeout;
4. retornar resultado;
5. ser registradas;
6. evitar execução em massa sem controles adicionais.

---

## 8. Defender e Firewall

A Central consulta postura de segurança, mas não oferece atalhos para desativar Defender ou Firewall.

Essa é uma decisão de design.

Se uma investigação exigir alteração de controle de segurança, ela deve seguir procedimento corporativo específico fora de um botão genérico da Central.

---

## 9. Sysinternals

A Central não baixa Sysinternals automaticamente.

Motivos:

- controle de versão;
- cadeia de confiança;
- ambientes sem acesso à internet;
- políticas de software homologado;
- evitar introdução silenciosa de executáveis administrativos.

O diretório deve ser preparado e validado pelo administrador.

---

## 10. Pacotes de diagnóstico

Relatórios podem conter:

- hostname;
- usuário;
- IP;
- software;
- eventos;
- caminhos de arquivo;
- serviços;
- processos;
- informações de domínio.

Trate esses arquivos como dados internos.

A pasta `reports/` é ignorada pelo Git, mas isso não substitui controle de acesso no filesystem.

---

## 11. Logs de auditoria

Logs administrativos devem ser úteis para rastreabilidade sem registrar segredos.

Idealmente registrar:

```text
data/hora
operador
host
ação
transporte
duração
sucesso/falha
erro sanitizado
```

Evite registrar conteúdo completo que possa conter dados sensíveis quando não houver necessidade operacional.

---

## 12. Injeção de comandos

Toda entrada do operador que seja inserida em PowerShell/CMD deve ser tratada como potencialmente perigosa.

Regras:

- validar tipos;
- limitar formatos;
- escapar strings;
- preferir parâmetros estruturados;
- evitar `shell=True`;
- evitar concatenar comandos arbitrários fornecidos pelo usuário;
- não criar “terminal remoto livre” como atalho de UI sem modelo de autorização/auditoria específico.

---

## 13. Princípio do menor privilégio

Mesmo sendo ferramenta N2, nem toda função precisa de máximo privilégio o tempo inteiro.

Futuras versões devem considerar:

- perfis de função;
- read-only vs remediação;
- autorização por módulo;
- restrição de ações em lote;
- confirmação reforçada para ações críticas.

---

## 14. Segurança de concorrência

Threads melhoram responsividade, mas não devem permitir duas remediações incompatíveis na mesma estação.

Exemplo a evitar:

```text
DISM RestoreHealth
+
Component Cleanup
+
reboot
```

simultaneamente.

A UI atual executa interações de forma sequencial para o operador, e o pool existe principalmente para não bloquear feedback visual.

---

## 15. Timeout não é cancelamento remoto

Esse é um ponto crítico.

Quando ocorre timeout:

```text
Central abandona Future
```

mas o processo remoto pode já estar executando.

Antes de repetir a ação:

- consulte processos/serviços;
- aguarde quando necessário;
- verifique logs/estado;
- não dispare a mesma manutenção pesada várias vezes.

---

## 16. Checklist de segurança antes de produção

```text
[ ] repositório sem segredos
[ ] settings revisado
[ ] PsExec homologado
[ ] Sysinternals homologado, se usado
[ ] WinRM configurado pela infraestrutura
[ ] operador autorizado
[ ] logs protegidos
[ ] reports protegidos
[ ] ações destrutivas testadas em bancada
[ ] baseline aprovado
[ ] estações piloto definidas
[ ] procedimento de rollback documentado
```

---

## 17. Incidente de segurança

Se houver suspeita de uso indevido da Central:

1. interrompa uso administrativo;
2. preserve logs;
3. identifique estação administrativa;
4. identifique hosts alvo;
5. revise histórico de ações;
6. revise credenciais/contextos utilizados;
7. acione segurança conforme procedimento corporativo;
8. não apague evidências como primeira ação.
