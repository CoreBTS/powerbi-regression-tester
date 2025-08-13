# runtime_load_adomd_dll.py
import os
import sys
import clr  # pythonnet
import System

# Always try to load from the unpacked _MEIPASS folder first
if hasattr(sys, '_MEIPASS'):
    dll_path = os.path.join(sys._MEIPASS, "Microsoft.AnalysisServices.AdomdClient.dll")
    if os.path.exists(dll_path):
        try:
            # Load directly from the DLL path to bypass GAC resolution
            clr.AddReferenceToFileAndPath(dll_path)
            print(f"[INFO] Loaded ADOMD.NET from embedded DLL: {dll_path}")
        except Exception as e:
            print(f"[ERROR] Failed to load ADOMD.NET from embedded DLL: {e}")
    else:
        print(f"[ERROR] Embedded ADOMD.NET DLL not found: {dll_path}")

# Optional: Verify loaded assembly version
try:
    asm = System.Reflection.Assembly.Load("Microsoft.AnalysisServices.AdomdClient")
    print(f"[INFO] Using ADOMD.NET version: {asm.GetName().Version}")
except Exception as e:
    print(f"[WARNING] Could not verify ADOMD.NET version: {e}")
