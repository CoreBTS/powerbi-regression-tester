# PowerBIRegressionTesterApp.spec
# Builds a single EXE that includes pyadomd and ADOMD.NET DLL.

from PyInstaller.utils.hooks import collect_submodules
import os
import sys

# Path to your ADOMD.NET DLL
ADOMD_DLL_PATH = r"C:\Program Files\Microsoft.NET\ADOMD.NET\160\Microsoft.AnalysisServices.AdomdClient.dll"
ADOMD_CONFIG_PATH = r"C:\Program Files\Microsoft.NET\ADOMD.NET\160\Microsoft.AnalysisServices.AdomdClient.replacements.config"

# Collect all pyadomd submodules and pythonnet's clr
hiddenimports = collect_submodules('pyadomd') + ['clr']

# Bundle the ADOMD.NET DLL and optional config into the EXE
binaries = [
    (ADOMD_DLL_PATH, '.'),  # DLL in root of _MEIxxxxx at runtime
]
if os.path.exists(ADOMD_CONFIG_PATH):
    binaries.append((ADOMD_CONFIG_PATH, '.'))

a = Analysis(
    ['PowerBIRegressionTesterApp.py'],  # your script
    pathex=[],
    binaries=binaries,
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['runtime_load_adomd_dll.py'],  # custom hook below
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

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
    console=False,  # same as --windowed
    disable_windowed_traceback=False
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PowerBIRegressionTesterApp'
)
