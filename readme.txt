####### Clone the project #######
git clone <repo-url>
cd myproject

####### Create the virtual environment #######
python -m venv venv

####### Activate the virtual environment #######
venv\Scripts\activate

####### Deactivate the virtual environment (if needed) #######
deactivate

####### Install requirements to the virtual environment #######
pip install -r requirements.txt
The requirements.txt was built using: pip freeze > requirements.txt

####### Pre-build Copy #######
Copy pyadomd.txt content into pyadomd.py in the venv where pyinstaller will run

####### Compile App #######
This app was built and tested with python 3.11.7
To create the .exe, normally you could run pyinstaller PowerBIRegressionTesterApp.py --onefile
However, since this app requires a copy of the ADOMD.dll a spec file is needed

####### Create single EXE #######
To create a single executable:
pyinstaller PowerBIRegressionTesterApp.spec --clean --noconfirm

####### Create EXE with supporting folder #######
To create an executable with the supporting files in a folder:
pyinstaller PowerBIRegressionTesterApp-Folder.spec --clean --noconfirm

####### Run App Manually #######
To start the app manually, open the PowerBIRegressionTesterApp.py file and run


####### To Do #######
More tooltips
Menus for config
