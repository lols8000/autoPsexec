# Changelog

## Unreleased — 2026-09-04

### Changed

- resultados estruturados em lista passam a ser apresentados como tabelas quando possível;
- inventário de drivers agora normaliza data, agrupa entradas idênticas e reduz registros vazios;
- view de drivers exibe Dispositivo, Fabricante, Versão, Data, Assinado, INF e Quantidade;
- parser JSON aceita payload válido mesmo quando o transporte adiciona texto antes/depois;
- PsExec remove mensagens operacionais sem valor em execuções bem-sucedidas;
- documentação atualizada para o comportamento v5 real, incluindo settings.local.json e fallback PsExec em ambientes sem WinRM.

### Fixed

- saída de inventário via PsExec não deve mais despejar Starting powershell.exe, exit code 0 ou CLIXML quando forem apenas ruído;
- JSON estruturado deixa de cair para texto bruto apenas por ruído de transporte;
- documentação antiga marcada como v3 foi promovida para v5.

### Tests

- teste de parser JSON com ruído PsExec;
- teste de limpeza de CLIXML em sucesso;
- teste garantindo preservação de erro real.

## 5.0.0 — 2026-09-03

### Added

- detecção de alvo local e transporte local;
- LocalTransport, WinRMTransport e PsExecTransport;
- TransportManager com cache e fallback;
- diagnóstico de conectividade e capabilities;
- SessionManager;
- JobManager e mutex por host;
- retry seletivo;
- progresso DISM;
- DiagnosticEngine e CorrelationEngine;
- nove playbooks;
- RemediationEngine com before/after;
- SQLite;
- diff de snapshots;
- baselines DEFAULT, DESKTOP, NOTEBOOK e TI;
- GLPI REST API opcional;
- correlation_id e sanitização;
- relatórios Markdown, JSON e TXT;
- console v5 e GUI Tkinter;
- atualização controlada;
- PyInstaller, Inno Setup e workflow de release.

### Security

- logger mascara senha, token, Authorization, API key, Bearer e credenciais.
