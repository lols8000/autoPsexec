# Troubleshooting — Central N2 Workstation 5.1.0

## Central não inicia

\`\`\`powershell
python --version
python .\central_n2\main.py --version
\`\`\`

Python 3.10+ é suportado pelo CI.

## Configuração inválida

Valide:

\`\`\`powershell
Test-Path .\central_n2\config\settings.json
\`\`\`

\`settings.local.json\` deve conter JSON válido e é aplicado por merge recursivo.

## UAC

Se a elevação falhar, execute o terminal autorizado como administrador e revise política/UAC. Não contorne o controle.

## Hostname não resolve

\`\`\`powershell
Resolve-DnsName PC023
ping PC023
\`\`\`

Estado esperado da Central: \`DNS_FAILED\`.

Prefira corrigir DNS a operar permanentemente por IP em domínio.

## Ping falha

Ping isolado não prova host offline. Teste serviços:

\`\`\`powershell
Test-NetConnection PC023 -Port 445
Test-NetConnection PC023 -Port 5985
\`\`\`

A Central avalia portas independentemente do ICMP.

## 5985 falha

Possíveis causas:

- WinRM parado;
- listener ausente;
- Firewall;
- LocalSubnet;
- ACL entre redes;
- GPO.

Se 445/ADMIN$/PsExec funcionarem, o host pode ficar \`READY_PSEXEC\`.

## WinRM por IP / CannotUseIPAddress

É tipicamente autenticação WinRM por IP, não reachability.

Use hostname/FQDN quando possível.

Evite \`TrustedHosts=*\`.

## Test-WSMan funciona, mas Invoke-Command falha

Na 5.1.0, o preflight já testa os dois. O host **não** deve ficar \`READY_WINRM\` se o \`Invoke-Command\` mínimo falhar.

Teste manual equivalente:

\`\`\`powershell
Test-WSMan PC023
Invoke-Command -ComputerName PC023 -ScriptBlock { 'CENTRAL_N2_WINRM_OK' }
\`\`\`

## PsExec não encontrado

\`\`\`powershell
Test-Path C:\Sysinternals\PsExec.exe
Test-Path C:\Windows\System32\PsExec.exe
Get-Command PsExec.exe -ErrorAction SilentlyContinue
\`\`\`

Copie apenas binário homologado para diretório controlado. A Central não o baixa.

## ADMIN$ funciona, mas PsExec não

ADMIN$ é requisito importante, mas não garante PsExec. Teste execução real:

\`\`\`powershell
C:\Sysinternals\PsExec.exe -accepteula -nobanner \\PC023 cmd.exe /d /c echo CENTRAL_N2_OK
\`\`\`

Revise EDR, SCM remoto, privilégio administrativo e política.

## AUTHENTICATION_FAILED

A rede responde, mas autenticação/autorização falhou em WinRM, ADMIN$ ou PsExec.

Não trate como falha de rede.

## NO_USABLE_TRANSPORT

O host é alcançável, porém nenhum transporte administrativo foi validado.

Use o menu Conectividade/Capabilities e separe cada camada.

## NETWORK_UNREACHABLE

Nenhum caminho administrativo conhecido respondeu. Revise rota, firewall, VLAN, host desligado e políticas.

## Resultado INDETERMINADO

Sintoma: a ação pode ter sido enviada, mas a comunicação caiu.

Conduta:

1. não repetir automaticamente;
2. consultar o estado final;
3. procurar evidência local/remota;
4. decidir conscientemente se precisa repetir.

O executor deve mostrar que o fallback foi suprimido.

## Timeout

Timeout da Central não prova que o processo remoto parou.

Antes de repetir DISM/SFC/install/reset, verifique processo, serviço ou log correspondente.

## Winget retorna falso sucesso

Na 5.1.0, install/upgrade/uninstall verificam \`$LASTEXITCODE\`. Se reaparecer falso sucesso, confirme a versão em execução.

## JSON não parseado após mutação

Quando uma mutação termina mas o retorno estruturado não pode ser validado, o resultado é marcado como indeterminado. Investigue antes de repetir.

## Spooler

Após remediação, o validador exige \`Status=Running\`.

UNKNOWN significa que a Central não conseguiu confirmar o estado final.

## Reset de Windows Update

A rotina registra os serviços originalmente ativos e tenta restaurá-los em \`finally\`. O validador exige confirmação de restauração.

## Compliance UNKNOWN

UNKNOWN significa métrica indisponível, não falha.

Exemplo: TPM/Secure Boot podem ser indisponíveis por firmware/cmdlet/permissão.

## Compliance N/A

N/A significa que o baseline não exige o controle. Não é erro de coleta.

## Bateria ausente

Normal em desktop ou hardware que não expõe WMI de bateria.

## Get-PhysicalDisk incompleto

RAID/controladores podem ocultar telemetria. Trate ausência de dados como limitação de evidência.

## Logs

Diretório:

\`\`\`text
central_n2\logs
\`\`\`

Não publique logs reais sem sanitização.

## SQLite

Se a Central reportar falha no \`quick_check\`, preserve o banco antes de qualquer tentativa de reparo.

Não altere \`user_version\` manualmente.

## Updater

Falhas possíveis:

- tag fora de SemVer;
- asset com nome inválido;
- URL não HTTPS;
- tamanho divergente;
- SHA-256 divergente;
- erro de rede.

Arquivo parcial deve ser removido e o destino final não deve ser promovido.

## Checklist de transporte

\`\`\`powershell
Resolve-DnsName PC023
Test-NetConnection PC023 -Port 445
Test-NetConnection PC023 -Port 5985
Test-Path \\PC023\ADMIN$
Test-Path C:\Sysinternals\PsExec.exe
Test-WSMan PC023
Invoke-Command -ComputerName PC023 -ScriptBlock { hostname }
\`\`\`
