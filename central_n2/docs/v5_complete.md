# Central N2 Workstation v5 — arquitetura consolidada

A v5 conclui o plano de evolução da Central N2 para uma plataforma modular de troubleshooting de estações Windows.

## Fluxo de execução

Alvo → HostIdentity → TransportManager → Local/WinRM/PsExec → Session + Connectivity + Capabilities → Snapshot → Findings → Correlação → Playbook → Remediação → Before/After → Relatório → GLPI opcional.

## Runtime e concorrência

O JobManager mantém um pool configurável. Leituras podem executar de forma concorrente; mutações são serializadas por host para evitar DISM, reset de Windows Update, reinicialização e outras ações pesadas concorrendo na mesma estação.

Timeout local não significa que um processo remoto já iniciado foi encerrado. Para processos explicitamente iniciados e rastreados existe RemoteProcessController, que devolve PID e permite cancelamento administrativo consciente.

## Sessões

SessionManager mantém um contexto lógico reutilizável por host com transporte selecionado, diagnóstico de conectividade e capabilities. O cache evita repetir preflight a cada ação. O PowerShell Remoting continua usando processos PowerShell independentes; a sessão lógica não deve ser confundida com um PSSession permanente do PowerShell.

## Capabilities

A Central detecta versão do PowerShell, Windows, Winget, Defender, BitLocker, TPM, Secure Boot, suporte a Get-PhysicalDisk, bateria e GLPI.

## Diagnóstico e playbooks

O motor separa fatos Finding de hipóteses Diagnosis. Os playbooks cobrem computador lento, rede, impressão, domínio/GPO, Windows Update, crash de aplicação, BSOD, disco cheio e GLPI Agent.

## Persistência e diff

SQLite em data/central_n2.db mantém histórico local. Banco, WAL e relatórios reais são ignorados pelo Git. O diff recursivo compara snapshots.

## Baselines

Perfis versionados: DEFAULT, DESKTOP, NOTEBOOK e TI. O compliance pode exigir BitLocker, TPM e Secure Boot conforme o perfil.

## GLPI

Credenciais e URL ficam exclusivamente em config/settings.local.json. A API é desabilitada por padrão. Quando habilitada, a Central cria acompanhamento em chamado usando o relatório gerado.

## Auditoria

Cada evento pode receber correlation_id. Antes de gravar, o logger mascara tokens, senhas, Authorization, Bearer, API keys e campos de credenciais.

## Interfaces

Console: python .\main.py
GUI: python .\main.py --gui

A GUI usa worker thread para não congelar a janela durante operações remotas.

## Distribuição

PyInstaller: pyinstaller .\CentralN2.spec
Instalador: iscc .\installer\CentralN2.iss
Tags v* acionam o workflow de release, que testa, compila, gera pacote portátil, gera instalador e publica ambos no GitHub Release.

## Atualização

A aplicação consulta somente o release mais recente e informa se existe nova versão. Download/instalação não é silencioso: atualização é deliberadamente controlada.

## Definition of Done

Novo recurso exige tratamento de erro, timeout, retorno visual, resultado estruturado, logging, teste e documentação.
