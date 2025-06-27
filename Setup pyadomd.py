import sys
import site
import os
import shutil
import importlib.util
import filecmp

"""
This script sets up the environment for using the Pyadomd library with ADOMD.NET.
Steps performed:
1. Ensures the ADOMD.NET path is present and adds it to the system path if necessary.
2. Prints the user site-packages directory.
3. Copies a custom .pth file to the user site-packages directory if it differs from the existing one.
4. Checks if the pyadomd module is installed, and if so, copies a custom pyadomd.txt file to the pyadomd module directory as pyadomd.py if it differs from the existing file.
5. Imports the Pyadomd class from the pyadomd module after setup.
Exits with an error message if required folders or modules are not found.
"""

# Ensure the ADOMD.NET path is in the system path prior to importing Pyadomd
# Adding to the path will usually work, but needs to be executed every time before importing Pyadomd.
# Instead of adding the path directly, 
# adomd_path = r'C:\Program Files\Microsoft.NET\ADOMD.NET\160'
# if not os.path.isdir(adomd_path):
#     print("Folder does not exist.")
#     sys.exit(1)
# # Check if the ADOMD.NET path is already in the system path
# if adomd_path not in path:
#     path.append(adomd_path)

# Adding a custom.pth file to the user site-packages directory will ensure that the ADOMD.NET path is included for all future sessions.
# This is a one-time setup, and the custom.pth file will persist across sessions.
# If this is not working, you can manually copy the custom.pth file to your site-packages directory folder
working_directory = os.getcwd()
src_file = os.path.join(working_directory, "TAA TD API", "custom.pth")  # Path to your source file in the subfolder
dst_sitepackage = next((p for p in site.getsitepackages() if "site-packages" in p.lower()), None)
if not dst_sitepackage:
    print("site-packages directory not found.")
    sys.exit(1)

dst_file = os.path.join(dst_sitepackage, "custom.pth")
if not os.path.exists(dst_file) or not filecmp.cmp(src_file, dst_file, shallow=False):
    shutil.copy(src_file, dst_file)

# Check if the pyadomd module is installed and get its folder
# If it is installed, copy the pyadomd.txt file to the module's folder
# If it is not installed, print an error message and exit, use pip install pyadomd to install it
# A common folder that it is installed in is C:\Users\username\AppData\Local\Programs\Python\Python311\Lib\site-packages\pyadomd
# The file pyadomd.txt should be copied to the same folder as pyadomd.py as it includes additional functionality
spec = importlib.util.find_spec("pyadomd")
if spec and spec.origin and len(spec.submodule_search_locations) > 0:
    pyadomd_folder = spec.submodule_search_locations[0]
    src_file = os.path.join(working_directory, "TAA TD API", "pyadomd.txt")  # Path to your source file in the subfolder
    dst_file = os.path.join(pyadomd_folder, "pyadomd.py")
    if not filecmp.cmp(src_file, dst_file, shallow=False):
        shutil.copy(src_file, dst_file)
else:
    print("pyadomd is not found, please install it using pip install pyadomd.")
    sys.exit(1)
