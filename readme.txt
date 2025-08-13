

Create the virtual environment using:
python -m venv venv

Activate the virtual environment using:
venv\Scripts\activate

Install requirements to the virtual environment using:
pip install -r requirements.txt

Deactivate the virtual environment using:
deactivate


To start the app, open the PowerBIRegressionTesterApp.py file and run

Copy pyadomd.txt content into pyadomd.py in the venv where pyinstaller will run
To create the .exe, run pyinstaller PowerBIRegressionTesterApp.py --onefile --windowed
pyinstaller PowerBIRegressionTesterApp.spec


To Do:
Exception Handling
py_installer
More tooltips
