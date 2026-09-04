# Central N2 Workstation v5

> Versão atual: **5.0.0** — plataforma de troubleshooting de estações Windows com transporte Local → WinRM → PsExec, diagnóstico correlacionado, playbooks, remediação guiada, histórico SQLite, relatórios e integração opcional com GLPI.

A Central N2 nasceu da evolução do projeto histórico **autoPsexec**. Os scripts legados da raiz permanecem como referência; a aplicação operacional atual está em **central_n2/**.

## O que a Central resolve

A Central organiza o atendimento N2 em um fluxo previsível:

~~~text
Selecionar estação
        ↓
Detectar transporte e capacidades
        ↓
Coletar evidências
        ↓
Diagnosticar / correlacionar
        ↓
Remediar com controle
        ↓
Validar antes/depois
        ↓
Registrar histórico / relatório / GLPI
~~~

Ela concentra rotinas que normalmente exigiriam PowerShell, CMD, Event Viewer, Device Manager, Windows Update, Sysinternals e ferramentas administrativas separadas.

## Transportes

A seleção é automática:

~~~text
Alvo local → LocalTransport
Alvo remoto com WinRM utilizável → WinRM
WinRM indisponível + PsExec disponível → PsExec
~~~

O WinRM **não é obrigatório**. Em ambientes onde a porta 5985 não está liberada, a Central pode operar por PsExec desde que:

- PsExec esteja instalado na estação administrativa;
- TCP 445/SMB esteja disponível;
- ADMIN$ esteja acessível;
- o operador tenha privilégio administrativo adequado.

Caminhos descobertos automaticamente incluem:

~~~text
C:\Windows\System32\PsExec.exe
C:\Sysinternals\PsExec.exe
~~~

## Saída legível

A v5 evita despejar JSON bruto para o técnico quando há dados estruturados. Listas como drivers, processos, serviços, impressoras e dispositivos são apresentadas em tabelas quando possível.

No transporte PsExec, mensagens operacionais sem valor para o técnico são filtradas em execuções bem-sucedidas, incluindo:

- Starting powershell.exe on ...
- powershell.exe exited on ... with error code 0
- envelope CLIXML gerado pelo Windows PowerShell/PsExec

Erros reais continuam preservados.

O inventário de drivers também:

- normaliza data;
- agrupa entradas idênticas;
- mostra assinatura como SIM/NÃO;
- reduz linhas vazias;
- inclui resumo de instâncias e drivers não assinados.

## Principais capacidades

- Saúde e compliance: CPU, RAM, disco, uptime, reboot pendente, Defender, Firewall, GLPI, BitLocker, TPM e Secure Boot.
- Performance: amostragem de CPU, memória, disco e rede, com processos dominantes.
- Reparo Windows: SFC, DISM, CHKDSK, Component Store e WMI.
- Hardware e drivers: PnP com erro, drivers, USB, rescan e exportação.
- Crashes/BSOD: BugCheck, Minidump, MEMORY.DMP e Application Error.
- Segurança: Defender, ameaças, Firewall, BitLocker, TPM, Secure Boot, RDP, SMBv1 e UAC.
- Rede: IP, gateway, DNS, DHCP, ARP, TCP, flush DNS e renovação DHCP.
- Usuários/perfis: sessões, administradores locais, perfis e tamanho.
- Software/GLPI Agent: inventário, Winget, status e inventário GLPI.
- Impressoras: inventário, fila e Spooler.
- Domínio/GPO: domínio, DC, secure channel, horário, gpresult e gpupdate.
- Armazenamento: volumes, Get-PhysicalDisk, saúde e bateria.
- Sysinternals: Autorunsc, ProcDump, Handle e Sigcheck.
- Playbooks: lentidão, rede, impressão, domínio/GPO, Windows Update, crash, BSOD, disco cheio e GLPI.
- Remediações guiadas: before/after e persistência.
- Histórico/Diff: SQLite e comparação de snapshots.
- Relatórios: Markdown, JSON e TXT.
- GLPI API opcional: envio do relatório ao chamado.
- Atualização controlada: consulta e download de releases, sem substituição silenciosa.

## Instalação rápida

~~~powershell
git clone https://github.com/lols8000/autoPsexec.git
cd autoPsexec
python .\central_n2\main.py
~~~

Para atualizar uma cópia existente:

~~~powershell
git checkout master
git pull
~~~

GUI opcional:

~~~powershell
python .\central_n2\main.py --gui
~~~

## Configuração

Configuração pública:

~~~text
central_n2\config\settings.json
~~~

Override local não versionado:

~~~text
central_n2\config\settings.local.json
~~~

O ConfigLoader faz merge recursivo do arquivo público com o local. Tokens, URLs internas, caminhos privados e outras configurações do ambiente devem ficar no arquivo local.

## Documentação

- central_n2/README.md — guia rápido
- central_n2/docs/v5_complete.md — visão consolidada v5
- central_n2/docs/architecture.md — arquitetura
- central_n2/docs/modules.md — catálogo de módulos
- central_n2/docs/configuration.md — configuração
- central_n2/docs/operations.md — operação N2
- central_n2/docs/security.md — segurança
- central_n2/docs/troubleshooting.md — troubleshooting da própria Central
- central_n2/CHANGELOG.md — histórico

## Testes e distribuição

~~~powershell
cd central_n2
python -m pip install pytest pyinstaller
python -m compileall .
python -m pytest -q
pyinstaller CentralN2.spec --clean --noconfirm
~~~

O CI Windows também valida build portátil e compilação do instalador Inno Setup.

## Princípio operacional

**EVIDÊNCIA → DIAGNÓSTICO → REMEDIAÇÃO → VALIDAÇÃO → REGISTRO**

A Central não existe para executar SFC/DISM/reset/reboot indiscriminadamente; ela existe para reduzir tentativa e erro e tornar o atendimento N2 rastreável.
