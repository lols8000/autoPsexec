from __future__ import annotations
import json,threading,tkinter as tk
from pathlib import Path
from core.session import SessionManager
from modules.health import HealthModule
from playbooks import PlaybookRunner,builtin_playbooks
from modules.performance import PerformanceModule
from modules.disk import DiskModule
from modules.system import SystemModule
from modules.startup import StartupModule
class CentralN2TkApp:
    def __init__(self,root,executor,settings_path:Path):
        self.root=root;self.executor=executor;self.sessions=SessionManager(executor);self.health=HealthModule(executor);self.playbooks=builtin_playbooks();self.runner=PlaybookRunner();self.performance=PerformanceModule(executor);self.disk=DiskModule(executor);self.system=SystemModule(executor);self.startup=StartupModule(executor)
        root.title("Central N2 Workstation");root.geometry("900x620");top=tk.Frame(root);top.pack(fill="x",padx=10,pady=10);tk.Label(top,text="Estação:").pack(side="left");self.host=tk.StringVar(value="localhost");tk.Entry(top,textvariable=self.host,width=35).pack(side="left",padx=5)
        for text,fn in (("Conectar",self.connect),("Saúde",self.check_health),("Playbook Lentidão",self.slow_playbook)):tk.Button(top,text=text,command=fn).pack(side="left",padx=4)
        self.status=tk.Label(root,text="Pronto",anchor="w");self.status.pack(fill="x",padx=10);self.output=tk.Text(root,wrap="word");self.output.pack(fill="both",expand=True,padx=10,pady=10)
    def _run(self,label,func):
        self.status.config(text=f"{label} — executando...")
        def worker():
            try:value=func();text=json.dumps(value,ensure_ascii=False,indent=2,default=str) if not isinstance(value,str) else value
            except Exception as exc:text=f"ERRO: {type(exc).__name__}: {exc}"
            self.root.after(0,lambda:self._finish(label,text))
        threading.Thread(target=worker,daemon=True).start()
    def _finish(self,label,text):self.output.delete("1.0","end");self.output.insert("end",text);self.status.config(text=f"{label} — concluído")
    def connect(self):self._run("Preflight",lambda:self.sessions.open(self.host.get().strip(),refresh=True))
    def check_health(self):
        h=self.host.get().strip()
        def f():
            r=self.health.snapshot(h);return r.data if r.success else r.stderr
        self._run("Saúde",f)
    def slow_playbook(self):
        h=self.host.get().strip();spec=self.playbooks["slow"];collectors={"health":self.health.snapshot,"performance":lambda x:self.performance.snapshot(x,8,1),"disk":self.disk.space,"processes":self.system.list_processes,"startup":self.startup.overview};self._run("Playbook Lentidão",lambda:self.runner.run(spec,h,collectors))
def run_gui(executor,settings_path:Path):root=tk.Tk();CentralN2TkApp(root,executor,settings_path);root.mainloop()
