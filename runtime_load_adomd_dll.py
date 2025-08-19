# runtime_load_adomd_dll.py
import os
import sys
import clr  # pythonnet
import System  # type: ignore
from sys import path
from System.Diagnostics import FileVersionInfo  # type: ignore

def log(msg):
    print(f"[RUNTIME HOOK] {msg}")

try:
    # exe_dir = os.path.dirname(sys.executable)
    # exe_dir = r"C:\Users\Cory.Cundy\source\repos\Power BI Regression Testing\dist\PowerBIRegressionTesterApp"
    # dll_dir = os.path.join(exe_dir, "_internal")
    # dll_path = os.path.join(dll_dir, "Microsoft.AnalysisServices.AdomdClient.dll")

    # if not os.path.exists(dll_path):
    #     raise FileNotFoundError(f"Embedded DLL not found at: {dll_path}")

    # asm = System.Reflection.Assembly.LoadFile(dll_path)

    # adomd_path = r'C:\Program Files\Microsoft.NET\ADOMD.NET\160'

    # if not os.path.isdir(dll_dir):
    #     print("Folder does not exist.")
    #     sys.exit(1)

    # # Check if the ADOMD.NET path is already in the system path
    # if dll_dir not in path:
    #     path.append(adomd_path)


    # if not hasattr(clr, "AddReferenceToFileAndPath"):
    #     raise AttributeError("pythonnet clr does not have AddReferenceToFileAndPath")

    # clr.AddReferenceToFileAndPath(dll_path)
    # log(f"Successfully loaded ADOMD.NET directly from: {dll_path}")

    # Verify version

    # Try loading the assembly by name
    log("\n=== Attempting to load ADOMD.NET assembly by name ===")
    clr.AddReference('Microsoft.AnalysisServices.AdomdClient')
    from Microsoft.AnalysisServices.AdomdClient import AdomdConnection, AdomdCommand # type: ignore
    log("✅ Successfully loaded ADOMD.NET assembly by name")

    target = None
    for asm in System.AppDomain.CurrentDomain.GetAssemblies():
        if asm.GetName().Name == "Microsoft.AnalysisServices.AdomdClient":
            target = asm
            break

    if not target:
        log("❌ ADOMD.NET assembly not found in loaded assemblies.")
        log("Ensure ADOMD.NET is installed and the path is correct.")
        log("https://learn.microsoft.com/en-us/analysis-services/client-libraries")
        log("dotnet add package Microsoft.AnalysisServices.AdomdClient --version 19.103.2")
    else:

        # Gather info
        asm_path = target.Location
        asm_ver = target.GetName().Version
        file_ver = FileVersionInfo.GetVersionInfo(asm_path).FileVersion
        in_gac = target.GlobalAssemblyCache

        # Output results
        log("✅ ADOMD.NET assembly detected")
        log(f"Assembly path   : {asm_path}")
        log(f"Assembly version: {asm_ver}")
        log(f"File version    : {file_ver}")
        log(f"Loaded from GAC : {in_gac}")



        # log(f"Successfully loaded ADOMD.NET from: {dll_path}")


        # log(f"Using ADOMD.NET Assembly version: {asm.GetName().Version}")
        # log(f"Using ADOMD.NET version: {asm.GetName().FullName}")
        # log(f"Using ADOMD.NET version: {System.Diagnostics.FileVersionInfo.GetVersionInfo(dll_path).FileVersion}")

except Exception as e:
    log(f"ERROR: Failed to load ADOMD.NET from packaged DLL: {e}")
    log(f"Type of exception: {type(e).__name__}: {type(e)}")
    #sys.exit(1)  # hard fail to ensure no GAC fallback
