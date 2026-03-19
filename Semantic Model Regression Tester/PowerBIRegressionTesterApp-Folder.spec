# -*- mode: python ; coding: utf-8 -*-
import os
import glob
from PyInstaller.utils.hooks import collect_submodules

# Collect all ADOMD.NET DLLs
ADOMD_DIR = r"C:\Program Files\Microsoft.NET\ADOMD.NET\160"
binaries = [(f, '.') for f in glob.glob(os.path.join(ADOMD_DIR, '*.dll'))]

# Collect all pyadomd submodules
# hidden_imports = collect_submodules('pyadomd') + ['clr']

# Optional runtime hook to load ADOMD.dll from temp extraction folder
# runtime_hooks=[r"runtime_load_adomd_dll.py"]

# ---------- ANALYSIS ----------
a = Analysis(
    ['PowerBIRegressionTesterApp.py'],
    pathex=[],
    binaries=binaries,
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

# ---------- PYZ ----------
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

# ---------- EXECUTABLE ----------
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PowerBIRegressionTesterApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=True,  # console enabled
    onefile=True,
    disable_windowed_traceback=False,
)

# COLLECT: gather everything into a single folder
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PowerBIRegressionTesterApp',
)
