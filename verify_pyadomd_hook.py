# verify_pyadomd_hook.py
import os
import sys
import traceback

try:
    import pyadomd
    print(f"[INFO] pyadomd imported successfully from: {pyadomd.__file__}")
except Exception as e:
    print(f"[ERROR] Failed to import pyadomd: {e}")
    traceback.print_exc()

# try:
#     import clr
#     dll_path = os.path.join(sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.dirname(sys.argv[0]),
#                             "Microsoft.AnalysisServices.AdomdClient.dll")
#     if os.path.exists(dll_path):
#         clr.AddReferenceToFileAndPath(dll_path)
#         print(f"[INFO] ADOMD.NET DLL loaded from: {dll_path}")
#     else:
#         print(f"[ERROR] ADOMD.NET DLL not found at: {dll_path}")
# except Exception as e:
#     print(f"[ERROR] Failed to load ADOMD.NET DLL: {e}")
#     traceback.print_exc()
