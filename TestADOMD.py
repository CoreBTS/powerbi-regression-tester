import sys
import clr  # from pythonnet
import System
from System.Diagnostics import FileVersionInfo

print("=== Before importing Pyadomd ===")
for a in System.AppDomain.CurrentDomain.GetAssemblies():
    if a.GetName().Name == "Microsoft.AnalysisServices.AdomdClient":
        print(f"Already loaded! Path: {a.Location}")
        break
else:
    print("AdomdClient is NOT loaded yet.")

# --- Import Pyadomd (will trigger ADOMD.NET load) ---
from pyadomd import Pyadomd

print("\n=== After importing Pyadomd ===")
for a in System.AppDomain.CurrentDomain.GetAssemblies():
    if a.GetName().Name == "Microsoft.AnalysisServices.AdomdClient":
        asm_path = a.Location
        file_ver = FileVersionInfo.GetVersionInfo(asm_path).FileVersion
        asm_ver = a.GetName().Version
        print(f"Assembly path   : {asm_path}")
        print(f"Assembly version: {asm_ver}")
        print(f"File version    : {file_ver}")
        break
else:
    print("AdomdClient STILL not loaded (unexpected).")
