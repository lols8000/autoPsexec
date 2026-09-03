# Troubleshooting — Central N2 Workstation v3

Este documento trata problemas da **própria Central N2** e de seus transportes.

## 1. A aplicação não inicia

### Sintoma

```text
'python' não é reconhecido
```

### Verificar

```powershell
python --version
py --version
```

É necessário Python 3.10+.

---

## 2. Configuração não encontrada

### Sintoma

```text
Arquivo de configuração não encontrado
```

### Verificar

```powershell
Test-Path .\central_n2\config\settings.json
```

Execute a partir da raiz correta do repositório ou preserve a estrutura de diretórios.

---

## 3. Falha ao elevar privilégio

### Sintoma

Mensagem de falha de UAC/ShellExecute.

### Verificar

- conta possui permissão administrativa;
- UAC não foi bloqueado por política;
- Python executável está acessível;
- terminal não está em contexto restrito.

Teste manual:

```powershell
Start-Process powershell -Verb RunAs
```

---

## 4. Host não resolve

### Sintoma

Hostname não é encontrado.

### Testes

```powershell
Resolve-DnsName PC023
ping PC023
```

Se IP funcionar e hostname não:

- revisar DNS;
- sufixo DNS;
- registro A/PTR;
- VPN/rede atual;
- cache DNS local.

Evite corrigir apenas usando IP para sempre; o DNS inconsistente deve ser tratado.

---

## 5. Ping falha

Ping falhar não prova necessariamente que o host está offline, porque ICMP pode ser bloqueado.

Valide outros serviços:

```powershell
Test-NetConnection PC023 -Port 5985
Test-NetConnection PC023 -Port 445
```

Considere política de firewall.

---

## 6. WinRM indisponível

### Teste

```powershell
Test-WSMan PC023
```

### No host alvo

```powershell
Get-Service WinRM
winrm enumerate winrm/config/listener
```

### Possíveis causas

- serviço parado;
- listener ausente;
- firewall;
- política de domínio;
- autenticação;
- DNS/SPN;
- perfil de rede;
- estação fora da rede corporativa.

### Importante

Não aplique configurações WinRM permissivas indiscriminadamente. Ajuste conforme política do ambiente.

---

## 7. WinRM funciona, mas comando falha

Verifique:

- privilégio do operador;
- cmdlet disponível no Windows alvo;
- versão/edição do Windows;
- módulo PowerShell instalado;
- comando não disponível em sessão remota;
- restrições de execução;
- Double Hop quando acesso a segundo recurso de rede é necessário.

---

## 8. PsExec não encontrado

### Teste

```powershell
Test-Path "C:\Windows\System32\PsExec.exe"
Get-Command PsExec.exe -ErrorAction SilentlyContinue
```

Ajuste `psexec_path` se necessário.

---

## 9. PsExec retorna Access Denied

Possíveis causas:

- conta sem administração local;
- UAC remoto;
- SMB/RPC bloqueado;
- `ADMIN$` indisponível;
- políticas de segurança;
- EDR bloqueando PsExec;
- estação fora do domínio/contexto esperado.

Testes:

```powershell
Test-NetConnection PC023 -Port 445
Test-Path \\PC023\ADMIN$
```

Não desabilite segurança globalmente para contornar o erro.

---

## 10. `ADMIN$` falha

Verifique:

- Server service;
- compartilhamentos administrativos;
- firewall;
- credenciais/contexto;
- política local/de domínio.

```powershell
net view \\PC023
```

---

## 11. A Central mostra heartbeat por muito tempo

Heartbeat indica que a UI está viva, mas a operação ainda não retornou.

Perguntas:

1. a operação é naturalmente longa, como DISM/SFC?
2. a estação está com CPU/disco saturado?
3. a rede está degradada?
4. o comando remoto está aguardando alguma condição?
5. o timeout configurado é razoável?

Não interrompa automaticamente só porque levou alguns minutos.

---

## 12. Timeout

### Sintoma

```text
✗ TIMEOUT
```

### Procedimento

1. registre qual operação estava rodando;
2. não repita imediatamente uma manutenção pesada;
3. verifique processos na estação;
4. verifique logs/eventos;
5. confirme se a operação remota continua;
6. só então decida repetir.

Timeout local não garante encerramento remoto.

---

## 13. UI parece parada e não há heartbeat

Isso pode indicar uma operação que não passou pelo `ResponsiveJobRunner` ou um erro anterior à execução do job.

Verifique:

- traceback no console;
- função chamada diretamente sem `execute()` na UI;
- input aguardando operador;
- erro de import;
- bloqueio durante inicialização.

Ao adicionar novos menus, toda operação longa deve passar pelo wrapper responsivo.

---

## 14. Winget não funciona remotamente

Winget depende fortemente de contexto.

Sob PsExec/SYSTEM ele pode:

- não estar no PATH;
- não possuir App Installer no contexto esperado;
- não enxergar fontes do usuário;
- não conseguir instalar pacote que exige sessão interativa.

Teste local no host:

```powershell
winget --info
Get-Command winget.exe
```

Se funcionar para usuário e não para SYSTEM, trate como limitação de contexto, não necessariamente bug da Central.

---

## 15. GLPI status funciona, instalação não

Verifique:

```json
"glpi": {
  "installer_source": "..."
}
```

O valor público vem vazio.

Valide acesso ao compartilhamento a partir do contexto utilizado pela execução remota.

Atenção a Double Hop quando WinRM precisa acessar um share de terceiro servidor.

---

## 16. Sysinternals aparece como ausente

Verifique:

```powershell
Get-ChildItem C:\Sysinternals
```

E configuração:

```json
"sysinternals_dir": "C:\\Sysinternals"
```

Nomes esperados incluem `Autorunsc.exe`, `ProcDump.exe`, `Handle.exe` e `Sigcheck.exe`.

---

## 17. ProcDump falha

Verifique:

- executável presente;
- processo existe;
- privilégio suficiente;
- caminho de saída permitido;
- espaço em disco;
- EDR não bloqueou criação do dump.

Dumps podem ser grandes e conter dados sensíveis em memória.

---

## 18. Get-Printer falha

Possíveis causas:

- módulo PrintManagement indisponível;
- Spooler parado;
- versão/edição Windows;
- sessão remota com restrições.

Teste:

```powershell
Get-Service Spooler
Get-Command Get-Printer
```

---

## 19. Get-BitLockerVolume falha

Possíveis causas:

- cmdlet não disponível;
- edição do Windows;
- recurso BitLocker ausente;
- privilégio insuficiente.

O módulo trata várias consultas como opcionais; ausência de dado não deve ser interpretada automaticamente como BitLocker desativado.

---

## 20. TPM/Secure Boot retornam vazio

Pode ocorrer em:

- hardware sem TPM;
- BIOS legado;
- VM;
- equipamento sem UEFI;
- cmdlet não suportado;
- restrição remota.

Considere `null` como “não foi possível obter” antes de concluir “desativado”.

---

## 21. Dados de bateria ausentes

Normal em:

- desktops;
- bateria não exposta via WMI;
- firmware limitado;
- equipamento sem driver ACPI adequado.

---

## 22. `Get-PhysicalDisk` não mostra saúde esperada

Alguns controladores RAID, drivers e dispositivos USB abstraem os dados.

O resultado representa o que o Windows Storage Management expõe, não telemetria SMART universal.

---

## 23. Security Event Log falha

Consultar log Security normalmente exige privilégio elevado.

Também pode ser limitado por:

- tamanho do log;
- auditoria desabilitada;
- filtro sem eventos;
- acesso remoto.

---

## 24. Testes falham

Execute:

```powershell
cd central_n2
python -m compileall .
python -m pytest -q
```

Para um teste específico:

```powershell
python -m pytest tests\test_v3_responsive.py -vv
```

### Antes de alterar código por causa de teste

Determine se:

- o teste está errado;
- o comportamento mudou intencionalmente;
- existe regressão real;
- é flutuação temporal/race condition.

---

## 25. Logs

Consulte:

```text
central_n2/logs/
```

Use logs para correlacionar:

- ação;
- host;
- transporte;
- duração;
- erro.

Não publique logs reais em issue/repositório público sem sanitização.

---

## 26. Pacotes de diagnóstico

Diretório esperado:

```text
central_n2/reports/diagnostics/
```

Se a geração falhar:

- validar permissão local;
- espaço em disco;
- caracteres inválidos;
- algum módulo remoto que não retornou;
- diretório de reports.

---

## 27. Checklist rápido de conectividade

```powershell
Resolve-DnsName PC023
Test-NetConnection PC023 -Port 5985
Test-NetConnection PC023 -Port 445
Test-WSMan PC023
Test-Path \\PC023\ADMIN$
```

Interprete cada teste separadamente; não reduza tudo a “máquina offline”.
