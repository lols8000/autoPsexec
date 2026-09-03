from __future__ import annotations
import json
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Callable
from core.baselines import BaselineRepository
from core.config import ConfigLoader,deep_merge
from core.jobs import JobManager,OperationClass
from core.result import CommandResult
from core.session import SessionManager
from core.updater import UpdateManager
from core.version import __version__
from diagnostics.engine import DiagnosticEngine
from diagnostics.correlation import CorrelationEngine
from integrations.glpi.client import GLPIClient,GLPIError
from modules.compliance import evaluate_compliance
from playbooks import PlaybookRunner,builtin_playbooks
from remediation import RemediationEngine,RemediationSpec
from reports import ReportExporter,SupportReportBuilder
from storage import CentralDatabase,diff_values
from ui.console_v3 import ConsoleUIV3
class ConsoleUIV5(ConsoleUIV3):
    def __init__(self,executor,settings_path:Path)->None:
        super().__init__(executor,settings_path);self.settings=ConfigLoader(settings_path).settings;runtime=self.settings.get("runtime",{})
        self.job_manager=JobManager(max_workers=int(runtime.get("max_workers",6)),heartbeat_seconds=float(self.settings.get("ui",{}).get("heartbeat_seconds",.2)));self.sessions=SessionManager(executor);self.baselines=BaselineRepository(settings_path.parent.parent/"baselines")
        persistence=self.settings.get("persistence",{});db_path=Path(persistence.get("database","data/central_n2.db"));db_path=db_path if db_path.is_absolute() else settings_path.parent.parent/db_path;self.db=CentralDatabase(db_path) if persistence.get("enabled",True) else None
        self.job_manager.set_observer(self.db.save_job if self.db else None)
        self.active_baseline_profile=str(self.settings.get("compliance",{}).get("profile","DEFAULT")).upper()
        self.engine=DiagnosticEngine();self.correlator=CorrelationEngine();self.playbook_runner=PlaybookRunner();self.playbooks=builtin_playbooks();self.remediation_engine=RemediationEngine();self.report_builder=SupportReportBuilder();self.report_exporter=ReportExporter(settings_path.parent.parent/"reports"/"support")
        cfg=self.settings.get("updates",{});self.updater=UpdateManager(cfg.get("repository","lols8000/autoPsexec"),__version__);self.update_dir=settings_path.parent.parent/"updates";self.last_diagnoses=[];self.last_playbook=None;self.last_remediation=None;self.last_report_path=None
    def run(self):
        try:
            while True:
                self.clear();print("╔════════════════════════════════════════════════════╗");print(f"║ CENTRAL N2 WORKSTATION — V{__version__:<24}║");print("╚════════════════════════════════════════════════════╝");transport=self.executor.select_transport(self.host) if self.host else "-";print(f"\nAlvo: {self.host or 'nenhum'} | Transporte: {transport}")
                lines=["[1] Selecionar estação","[2] Saúde / Compliance","[3] Performance","[4] Reparo do Windows","[5] Hardware / Drivers / Dispositivos","[6] Inicialização / Tarefas","[7] Crashes / BSOD","[8] Segurança","[9] Rede","[10] Usuários / Perfis","[11] Software / GLPI Agent","[12] Impressoras","[13] Domínio / GPO","[14] Disco / Armazenamento / Bateria","[15] Ferramentas avançadas","[16] Sysinternals","[17] Pacote de diagnóstico","[18] Energia / Processos / Serviços","[19] Conectividade / Capabilities","[20] Assistente N2 / Playbooks","[21] Histórico / Diff","[22] Gerar relatório","[23] Jobs","[24] Atualização da Central","[25] GLPI API","[26] Remediações guiadas","[27] Perfil / Baseline","[0] Sair"]
                for line in lines:print(line)
                op=input("\nOpção: ").strip();handlers={"1":self.select_host,"2":self.menu_health,"3":self.menu_performance,"4":self.menu_repair,"5":self.menu_devices,"6":self.menu_startup_tasks,"7":self.menu_crashes,"8":self.menu_security,"9":self.menu_network,"10":self.menu_users,"11":self.menu_software_glpi,"12":self.menu_printers,"13":self.menu_domain,"14":self.menu_storage,"15":self.menu_tools,"16":self.menu_sysinternals,"17":self.collect_diagnostic,"18":self.menu_system,"19":self.menu_connectivity,"20":self.menu_playbooks,"21":self.menu_history,"22":self.menu_report,"23":self.menu_jobs,"24":self.menu_update,"25":self.menu_glpi_api,"26":self.menu_remediations,"27":self.menu_baseline}
                if op=="0":return
                if op in handlers:handlers[op]()
        finally:self.jobs.shutdown();self.job_manager.shutdown()
    def select_host(self):
        self.clear();host=input("Hostname ou IP da estação: ").strip()
        if not host:return
        try:s=self.jobs.run("Abrindo sessão lógica",lambda:self.sessions.open(host,refresh=True),timeout=180)
        except Exception as exc:print(f"\n✗ Falha no preflight: {exc}");self.pause();return
        self.host=host;print(f"\n✓ Transporte selecionado: {s.transport}");print(f"Local: {'SIM' if s.connectivity.get('is_local') else 'NÃO'} | DNS: {s.connectivity.get('dns')} | WinRM: {s.connectivity.get('winrm')} | ADMIN$: {s.connectivity.get('admin_share')}");print(f"Diagnóstico: {s.connectivity.get('diagnosis')}")
        r=self.execute("Snapshot inicial de saúde",lambda:self.health.snapshot(host),timeout=120)
        if r and r.success and isinstance(r.data,dict):self.health_snapshot=r.data;self.show_health(r.data);self._persist_health(r.data)
        self.pause()
    def _baseline(self):
        cfg=self.settings.get("compliance",{})
        effective=dict(cfg)
        effective["profile"]=self.active_baseline_profile
        return deep_merge(self.baselines.load(self.active_baseline_profile),effective)

    @staticmethod
    def _classify_operation(label:str)->OperationClass:
        value=label.lower()
        if any(key in value for key in ("deslig","reinicialização","reiniciar estação")):
            return OperationClass.DISRUPTIVE
        if any(key in value for key in ("dism","sfc","chkdsk","component store","reset", "limpeza")):
            return OperationClass.HEAVY_WRITE
        if any(key in value for key in ("gpupdate","flush dns","renovar dhcp","spooler","forçando inventário","enviando mensagem","reexaminando","exportando drivers")):
            return OperationClass.LIGHT_WRITE
        return OperationClass.READ_ONLY

    def execute(
        self,
        label:str,
        func:Callable[[],CommandResult],
        *,
        timeout:int|None=None,
        operation_class:OperationClass|None=None,
    )->CommandResult|None:
        print(f"\n▶ {label}")
        op_class=operation_class or self._classify_operation(label)
        target=self.host or "local-ui"
        try:
            result=self.job_manager.run_sync(
                target,
                label,
                func,
                operation_class=op_class,
                timeout=timeout or self.long_timeout,
            )
        except TimeoutError as exc:
            print(f"\n✗ TIMEOUT: {exc}")
            return None
        except Exception as exc:
            print(f"\n✗ ERRO NÃO TRATADO: {type(exc).__name__}: {exc}")
            return None
        self.show_result(result)
        return result
    def _persist_health(self,data):
        if self.db:self.db.save_snapshot(self.host,data,kind="health")
        findings=self.engine.evaluate(data);self.last_diagnoses=self.correlator.correlate(findings)
        if self.db:
            for f in findings:self.db.save_finding(self.host,f.id,f.severity.value,asdict(f))
    def menu_health(self):
        if not self.require_host():return
        self.clear();r=self.execute("Coletando saúde e compliance",lambda:self.health.snapshot(self.host),timeout=120)
        if r and r.success and isinstance(r.data,dict):
            self.health_snapshot=r.data;self.show_health(r.data);self._persist_health(r.data);comp=evaluate_compliance(r.data,self._baseline());print(f"\nCompliance: {comp['score']}/100 ({comp['compliant']}/{comp['total']})")
            for i in comp["items"]:print(f" {'✓' if i['compliant'] else '✗'} {i['label']}: {i['actual']} | esperado {i['expected']}")
            for d in self.last_diagnoses:print(f" - {d.title} | confiança {d.confidence}: {d.rationale}")
        self.pause()
    def menu_repair(self):
        if not self.require_host():return
        self.clear();print("REPARO DO WINDOWS\n1 - SFC\n2 - DISM CheckHealth\n3 - DISM ScanHealth\n4 - DISM RestoreHealth (progresso real)\n5 - Analisar Component Store\n6 - Limpar Component Store\n7 - CHKDSK online\n8 - Verificar WMI\n0 - Voltar");op=input("Opção: ").strip()
        if op=="4":
            if not self.confirm(f"Executar DISM RestoreHealth em {self.host}?"):return
            def progress(ev):
                if ev.percent is not None:print(f"\rDISM RestoreHealth [{ev.percent:6.2f}%] {ev.message[:60]:60}",end="",flush=True)
            try:r=self.job_manager.run_sync(self.host,"DISM RestoreHealth",lambda:self.repair.dism_restorehealth_progress(self.host,progress),operation_class=OperationClass.HEAVY_WRITE,timeout=3600,on_tick=lambda _:None);print();self.show_result(r)
            except Exception as exc:print(f"\n✗ {exc}")
            self.pause();return
        actions={"1":("SFC /scannow",self.repair.sfc_scan),"2":("DISM CheckHealth",self.repair.dism_checkhealth),"3":("DISM ScanHealth",self.repair.dism_scanhealth),"5":("Analisar Component Store",self.repair.component_store),"6":("Limpar Component Store",self.repair.component_cleanup),"7":("CHKDSK /scan",self.repair.chkdsk_scan),"8":("Verificar WMI",self.repair.repository_consistency)};item=actions.get(op)
        if not item:return
        if op=="6" and not self.confirm(f"Executar {item[0]}?"):return
        self.execute(item[0],lambda:item[1](self.host),timeout=3600);self.pause()
    def menu_connectivity(self):
        if not self.require_host():return
        self.clear()
        try:s=self.jobs.run("Atualizando conectividade/capabilities",lambda:self.sessions.open(self.host,refresh=True),timeout=180);print(json.dumps({"transport":s.transport,"connectivity":s.connectivity,"capabilities":s.capabilities},indent=2,ensure_ascii=False,default=str))
        except Exception as exc:print(exc)
        self.pause()
    def _collectors(self):
        return {"health":self.health.snapshot,"performance":lambda h:self.performance.snapshot(h,8,1),"disk":self.disk.space,"processes":self.system.list_processes,"startup":self.startup.overview,"network":self.network.ip_configuration,"adapters":self.network.adapters,"connections":self.network.connections,"proxy":self.tools.proxy,"printers":self.printers.list_printers,"print_queue":self.printers.queue,"services":self.system.list_services,"domain":self.domain.status,"gpresult":self.domain.gpresult,"updates":self.updates.status,"app_crashes":self.crashes.app_crashes,"bsod":self.crashes.bsod_history,"devices":self.devices.problem_devices,"profiles":self.disk.profile_sizes,"cleanup_estimate":self.disk.cleanup_estimate,"glpi":self.glpi.status,"glpi_log":self.glpi.recent_log}
    def menu_playbooks(self):
        if not self.require_host():return
        self.clear();keys=list(self.playbooks);print("ASSISTENTE N2 / PLAYBOOKS")
        for i,k in enumerate(keys,1):print(f"{i} - {self.playbooks[k].title}")
        op=input("Opção (0 volta): ").strip()
        if op=="0" or not op.isdigit() or not(1<=int(op)<=len(keys)):return
        spec=self.playbooks[keys[int(op)-1]]
        try:execution=self.job_manager.run_sync(self.host,f"Playbook {spec.title}",lambda:self.playbook_runner.run(spec,self.host,self._collectors(),on_step=lambda i,t,s:print(f"[{i}/{t}] {s.label}...")),operation_class=OperationClass.READ_ONLY,timeout=self.long_timeout,on_tick=lambda _:None)
        except Exception as exc:print(f"✗ {exc}");self.pause();return
        self.last_playbook=execution;merged={}
        for st in execution.steps:
            if st.get("success") and isinstance(st.get("data"),dict):merged.update(st["data"])
            print(f"{'✓' if st.get('success') else '✗'} {st['label']} [{st.get('transport','-')}] {st.get('error') or ''}")
        findings=self.engine.evaluate(merged);self.last_diagnoses=self.correlator.correlate(findings)
        if self.db:self.db.save_snapshot(self.host,asdict(execution),kind=f"playbook:{spec.key}")
        for d in self.last_diagnoses:print(f"- {d.title} ({d.confidence}): {d.rationale}")
        self.pause()
    def menu_history(self):
        if not self.require_host():return
        self.clear()
        if not self.db:print("Persistência desabilitada.");self.pause();return
        for r in self.db.recent_snapshots(self.host,limit=10):print(f"#{r['id']} {r['created_at']} {r['kind']}")
        print("\nDIFF DOS DOIS ÚLTIMOS HEALTH:")
        changes=self.db.diff_latest(self.host,kind="health")
        if not changes:print("Sem dois snapshots comparáveis.")
        for c in changes[:100]:print(f"- {c['path']}: {c['before']} -> {c['after']} ({c['type']})")
        self.pause()
    def menu_report(self):
        if not self.require_host():return
        self.clear();problem=input("Problema/resumo do chamado: ").strip();diag="; ".join(f"{d.title} ({d.confidence})" for d in self.last_diagnoses) or "Sem diagnóstico correlacionado registrado.";actions=[x["label"] for x in self.last_playbook.steps if x.get("success")] if self.last_playbook else []
        if self.last_remediation:
            actions.append(f"Remediação: {self.last_remediation.spec.title} — {'sucesso' if self.last_remediation.command_result.success else 'falha'}")
        validation=self.last_remediation.after if self.last_remediation and self.last_remediation.after is not None else self.health_snapshot
        report=self.report_builder.build(host=self.host,user=(self.health_snapshot or {}).get("User"),problem=problem,diagnosis=diag,actions=actions,validation=validation,result="Diagnóstico/atendimento registrado");stamp=datetime.now().strftime("%Y%m%d_%H%M%S");path=self.report_exporter.export(report,fmt="markdown",stem=f"{self.host}_{stamp}");self.report_exporter.export(report,fmt="json",stem=f"{self.host}_{stamp}");self.last_report_path=path
        if self.db:self.db.save_report(self.host,"markdown",path.read_text(encoding="utf-8"),path=str(path))
        print(f"✓ Relatório: {path}");self.pause()
    def menu_jobs(self):
        self.clear()
        print("JOBS DA SESSÃO")
        records=self.job_manager.list_records()[-30:]
        if not records:
            print("Nenhum job nesta sessão.")
        for record in records:
            print(f"{record.job_id} | {record.host} | {record.state.value} | {record.operation_class.value} | {record.elapsed_seconds:.1f}s | {record.label}")

        if self.db:
            print("\nÚLTIMOS JOBS PERSISTIDOS")
            for item in self.db.recent_jobs(self.host,limit=15):
                print(f"{item['job_id']} | {item['host']} | {item['state']} | {item['label']}")
        self.pause()
    def menu_update(self):
        self.clear()
        print(f"Versão atual: {__version__}")
        try:
            info=self.jobs.run("Consultando releases",self.updater.check_latest,timeout=30)
        except Exception as exc:
            print(f"Não foi possível consultar atualização: {exc}")
            self.pause()
            return

        print(f"Última versão: {info.latest} | Atualização disponível: {'SIM' if info.update_available else 'NÃO'}")
        if not info.update_available or not info.assets:
            self.pause()
            return

        print("\nArtefatos disponíveis:")
        for index,asset in enumerate(info.assets,1):
            print(f"{index} - {asset.get('name','sem nome')}")

        choice=input("Número para baixar (0 cancela): ").strip()
        if choice=="0" or not choice.isdigit() or not (1<=int(choice)<=len(info.assets)):
            return

        asset=info.assets[int(choice)-1]
        if not self.confirm(f"Baixar {asset.get('name','artefato')} para a pasta updates?"):
            return
        try:
            path=self.jobs.run(
                "Baixando atualização",
                lambda:self.updater.download_asset(asset,self.update_dir),
                timeout=900,
            )
            print(f"✓ Download concluído: {path}")
            print("A instalação permanece manual/controlada; a Central não se substitui silenciosamente.")
        except Exception as exc:
            print(f"✗ Falha no download: {exc}")
        self.pause()
    def menu_glpi_api(self):
        self.clear();cfg=self.settings.get("glpi_api",{})
        if not cfg.get("enabled"):print("GLPI API desabilitada. Configure somente em settings.local.json.");self.pause();return
        ticket=input("ID do chamado GLPI: ").strip()
        if not ticket.isdigit():return
        if not self.last_report_path or not Path(self.last_report_path).exists():print("Gere um relatório antes de enviar ao GLPI.");self.pause();return
        client=GLPIClient(cfg.get("base_url",""),cfg.get("app_token",""),cfg.get("user_token",""))
        try:client.add_ticket_followup(int(ticket),Path(self.last_report_path).read_text(encoding="utf-8"));print("✓ Acompanhamento enviado ao GLPI.")
        except GLPIError as exc:print(f"✗ GLPI: {exc}")
        finally:
            try:client.kill_session()
            except Exception:pass
        self.pause()

    def _health_payload(self,host:str):
        result=self.health.snapshot(host)
        if result.success and isinstance(result.data,dict):
            return result.data
        return {"error":result.stderr or "snapshot indisponível"}

    def menu_remediations(self):
        if not self.require_host():
            return
        self.clear()
        print("REMEDIAÇÕES GUIADAS")
        print("1 - Limpeza segura de temporários")
        print("2 - Reiniciar Spooler")
        print("3 - Resetar componentes do Windows Update")
        print("4 - GPUpdate /force")
        print("0 - Voltar")
        op=input("Opção: ").strip()

        specs={
            "1":(
                RemediationSpec("safe_cleanup","Limpeza segura de temporários","médio",True,False,False,"Não há rollback automático para temporários removidos."),
                self.disk.cleanup_safe,
                OperationClass.HEAVY_WRITE,
            ),
            "2":(
                RemediationSpec("restart_spooler","Reiniciar Spooler","baixo",True,False,False),
                self.printers.restart_spooler,
                OperationClass.LIGHT_WRITE,
            ),
            "3":(
                RemediationSpec("reset_windows_update","Resetar componentes do Windows Update","alto",True,False,False,"SoftwareDistribution/catroot2 são renomeados com timestamp."),
                self.updates.reset_components,
                OperationClass.HEAVY_WRITE,
            ),
            "4":(
                RemediationSpec("gpupdate_force","GPUpdate /force","baixo",True,False,False),
                self.system.gpupdate,
                OperationClass.LIGHT_WRITE,
            ),
        }
        item=specs.get(op)
        if not item:
            return
        spec,action,operation_class=item
        print(f"\nAção: {spec.title}")
        print(f"Impacto: {spec.impact.upper()} | Reboot esperado: {'SIM' if spec.requires_reboot else 'NÃO'}")
        if spec.rollback:
            print(f"Rollback/observação: {spec.rollback}")
        if spec.requires_confirmation and not self.confirm(f"Executar {spec.title} em {self.host}?"):
            return

        try:
            remediation=self.job_manager.run_sync(
                self.host,
                f"Remediação: {spec.title}",
                lambda:self.remediation_engine.execute(
                    self.host,
                    spec,
                    action,
                    snapshotter=self._health_payload,
                ),
                operation_class=operation_class,
                timeout=self.long_timeout,
            )
        except Exception as exc:
            print(f"✗ Remediação falhou: {type(exc).__name__}: {exc}")
            self.pause()
            return

        self.last_remediation=remediation
        result=remediation.command_result
        self.show_result(result)
        changes=diff_values(remediation.before,remediation.after) if remediation.after is not None else []
        print("\nANTES / DEPOIS")
        if changes:
            for change in changes[:50]:
                print(f"- {change['path']}: {change['before']} -> {change['after']}")
        else:
            print("Nenhuma diferença do snapshot de saúde foi detectada.")

        if self.db:
            self.db.save_remediation(
                self.host,
                spec.key,
                result.success,
                asdict(remediation),
            )
            if isinstance(remediation.after,dict):
                self.db.save_snapshot(self.host,remediation.after,kind=f"after:{spec.key}")
        self.pause()

    def menu_baseline(self):
        self.clear()
        profiles=self.baselines.available()
        print(f"Perfil ativo: {self.active_baseline_profile}")
        for index,profile in enumerate(profiles,1):
            print(f"{index} - {profile}")
        print("0 - Voltar")
        choice=input("Perfil: ").strip()
        if choice=="0" or not choice.isdigit() or not (1<=int(choice)<=len(profiles)):
            return
        self.active_baseline_profile=profiles[int(choice)-1]
        print(f"✓ Baseline ativo nesta sessão: {self.active_baseline_profile}")
        self.pause()
