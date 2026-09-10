from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules

root = Path(SPECPATH)

datas = [
    (str(root / "config" / "settings.json"), "config"),
]
for baseline in (root / "baselines").glob("*.json"):
    datas.append((str(baseline), "baselines"))

hiddenimports = collect_submodules("tkinter")

analysis = Analysis(
    ["main.py"],
    pathex=[str(root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(analysis.pure)

exe = EXE(
    pyz,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="CentralN2",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    version=str(root / "version_info.txt"),
)

coll = COLLECT(
    exe,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=False,
    name="CentralN2",
)
