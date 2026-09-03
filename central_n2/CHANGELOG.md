# Changelog
## 5.0.0 — 2026-09-03
### Added
- detecção de alvo local e transporte local;
- camadas LocalTransport, WinRMTransport e PsExecTransport;
- TransportManager com cache e fallback seguro apenas para falha de transporte;
- diagnóstico de conectividade e capability detection;
- SessionManager para contexto lógico reutilizável;
- JobManager, classificação de operações e mutex por host;
- retry seletivo;
- progresso percentual para DISM RestoreHealth quando o comando fornece percentual;
- DiagnosticEngine, findings e CorrelationEngine;
- nove playbooks N2;
- RemediationEngine com snapshot before/after;
- SQLite para snapshots, findings, remediações, jobs e relatórios;
- diff recursivo de snapshots;
- baselines DEFAULT, DESKTOP, NOTEBOOK e TI;
- cliente GLPI REST API opcional;
- auditoria com correlation_id e sanitização de segredos;
- relatórios Markdown, JSON e TXT;
- console v5 e GUI Tkinter opcional;
- verificação controlada de releases;
- PyInstaller, Inno Setup e workflow de release.
### Changed
- health/compliance incluem BitLocker, TPM e Secure Boot;
- settings.local.json é o local recomendado para caminhos e tokens internos;
- execução local não exige WinRM quando o hostname selecionado é a própria estação.
### Security
- logger mascara senha, token, Authorization, API key, Bearer e credenciais antes de persistir.
