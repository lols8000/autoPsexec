# Troubleshooting — Central N2 Workstation v5

## Central não inicia

Valide Python 3.10+ ou use pacote distribuído.

~~~powershell
python --version
~~~

## Configuração não encontrada

~~~powershell
Test-Path .\central_n2\config\settings.json
~~~

## UAC

~~~powershell
Start-Process powershell -Verb RunAs
~~~

## Hostname não resolve

~~~powershell
Resolve-DnsName PC023
ping PC023
~~~

Prefira corrigir DNS a operar permanentemente por IP.

## Ping falha

Ping não prova sozinho host offline. Teste serviços.

~~~powershell
Test-NetConnection PC023 -Port 5985
Test-NetConnection PC023 -Port 445
~~~

## WinRM 5985 falha

Se:

~~~text
PingSucceeded: True
TcpTestSucceeded: False
RemotePort: 5985
~~~

a conexão não chegou ao serviço WinRM.

Possíveis causas:

- WinRM parado;
- listener ausente;
- firewall do endpoint;
- regra restrita a LocalSubnet;
- ACL/firewall entre redes;
- política de domínio.

No destino:

~~~powershell
Get-Service WinRM
winrm enumerate winrm/config/listener
Get-NetConnectionProfile
~~~

## WinRM por IP retorna CannotUseIPAddress

Isso é problema de autenticação WinRM por IP, não de reachability.

Em domínio, use hostname/FQDN.

Evite TrustedHosts=*.

## Test-WSMan funciona, Invoke-Command falha

Test-WSMan prova listener WSMan, não necessariamente autenticação PowerShell Remoting completa.

Teste:

~~~powershell
Invoke-Command -ComputerName PC023 -ScriptBlock { hostname }
~~~

## PsExec não encontrado

~~~powershell
Test-Path C:\Sysinternals\PsExec.exe
Test-Path C:\Windows\System32\PsExec.exe
Get-Command PsExec.exe -ErrorAction SilentlyContinue
~~~

## Instalar PsExec na estação administrativa

A Central não instala automaticamente. Use binário homologado da Microsoft Sysinternals e coloque em diretório controlado, por exemplo:

~~~text
C:\Sysinternals\PsExec.exe
~~~

Feche e reabra a Central depois de copiar o executável, pois o executor descobre PsExec na inicialização.

## Validar PsExec

~~~powershell
Test-NetConnection PC023 -Port 445
dir \\PC023\ADMIN$
C:\Sysinternals\PsExec.exe -accepteula \\PC023 hostname
~~~

## ADMIN$ funciona e WinRM não

Isso é um cenário suportado.

~~~text
445/ADMIN$ OK + PsExec disponível → fallback PsExec
~~~

Não é obrigatório abrir 5985 somente para a Central funcionar.

## PsExec Access Denied

Verifique:

- conta administrativa;
- ADMIN$;
- UAC remoto;
- política/EDR;
- contexto de domínio.

Não desative segurança globalmente.

## Saída mostra Starting powershell.exe / CLIXML

No master atual, isso deve ser filtrado em execuções PsExec bem-sucedidas.

Se reaparecer:

1. confirme git pull;
2. reinicie a Central;
3. confirme versão/commit;
4. preserve a saída em caso de falha real.

## JSON bruto em listas

A UI atual tenta renderizar listas de dicionários como tabela. Se JSON bruto aparecer, pode ser:

- estrutura complexa;
- parser não encontrou JSON válido;
- retorno textual do comando;
- execução em versão antiga.

## DriverDate aparece como /Date(...)/

O módulo de drivers atual normaliza data para yyyy-MM-dd antes de serializar.

Se aparecer formato legado, atualize o master e reinicie.

## Winget não funciona via PsExec

Winget pode depender do perfil do usuário/App Installer e não existir sob SYSTEM.

## GLPI

Se status funciona e instalação não, valide installer_source no settings.local.json e acesso ao recurso de origem.

## Sysinternals ausente

~~~powershell
Get-ChildItem C:\Sysinternals
~~~

## Timeout

Não repita uma remediação pesada imediatamente. Timeout local não prova término remoto.

## Bateria ausente

Normal em desktop ou firmware que não expõe dados.

## Get-PhysicalDisk incompleto

Controladores podem esconder telemetria.

## Logs

Consulte central_n2/logs. Não publique logs reais sem sanitização.

## Checklist rápido

~~~powershell
Resolve-DnsName PC023
Test-NetConnection PC023 -Port 5985
Test-NetConnection PC023 -Port 445
Test-Path \\PC023\ADMIN$
Test-Path C:\Sysinternals\PsExec.exe
Test-WSMan PC023
~~~

Interprete cada camada separadamente.
