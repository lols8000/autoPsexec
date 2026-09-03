from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules
root=Path(SPECPATH)
datas=[(str(root/"config"/"settings.json"),"config")]
for p in (root/"baselines").glob("*.json"):
    datas.append((str(p),"baselines"))
hiddenimports=collect_submodules("tkinter")
a=Analysis(["main.py"],pathex=[str(root)],binaries=[],datas=datas,hiddenimports=hiddenimports,hookspath=[],hooksconfig={},runtime_hooks=[],excludes=[],noarchive=False)
pyz=PYZ(a.pure)
exe=EXE(pyz,a.scripts,[],exclude_binaries=True,name="CentralN2",debug=False,bootloader_ignore_signals=False,strip=False,upx=True,console=True,disable_windowed_traceback=False)
coll=COLLECT(exe,a.binaries,a.datas,strip=False,upx=True,name="CentralN2")
