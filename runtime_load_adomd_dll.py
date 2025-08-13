# runtime_load_adomd_dll.py
import os
import sys
import clr  # pythonnet
import System

def log(msg):
    print(f"[RUNTIME HOOK] {msg}")

try:
    exe_dir = os.path.dirname(sys.executable)
    dll_dir = os.path.join(exe_dir, "_internal")
    dll_path = os.path.join(dll_dir, "Microsoft.AnalysisServices.AdomdClient.dll")

    if not os.path.exists(dll_path):
        raise FileNotFoundError(f"Embedded DLL not found at: {dll_path}")

    # if not hasattr(clr, "AddReferenceToFileAndPath"):
    #     raise AttributeError("pythonnet clr does not have AddReferenceToFileAndPath")

    # clr.AddReferenceToFileAndPath(dll_path)
    # log(f"Successfully loaded ADOMD.NET directly from: {dll_path}")

    # Verify version
    asm = System.Reflection.Assembly.LoadFile(dll_path)
    log(f"Using ADOMD.NET version: {asm.GetName().Version}")
    log(f"Using ADOMD.NET version: {asm.GetName().FullName}")

except Exception as e:
    log(f"ERROR: Failed to load ADOMD.NET from packaged DLL: {e}")
    #sys.exit(1)  # hard fail to ensure no GAC fallback
