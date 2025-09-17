# Power BI Regression Tester

The Power BI Regression Tester application will allow you to test changes made to a model to see if the change made a difference. This can be useful when trying to optimze DAX, change settings such as Filter Value Behavior switch or changing a relationship.

The application requires have access to two models that can be compared and input a set of Power BI DAX query files either frmo Power BI Performance Analyzer or a DAX Studio Trace. 

These queries can then be run on both models to see if there are differences. If there are differences, the individual queries can be run to see what the differences are both at the row level and column level.

## Setup
    1. Clone the project
    2. git clone https://core-bts-dadp-sandbox@dev.azure.com/core-bts-dadp-sandbox/Tabular_DAX_Sandbox/_git/Tabular_DAX_Sandbox
    3. cd myproject

## Create and activate the virtual environment
    1. python -m venv venv
    2. venv\Scripts\activate

    Deactivate the virtual environment using (deactivate)

## Install requirements to the virtual environment
    1. pip install -r requirements.txt
    2. The requirements.txt was built using: pip freeze > requirements.txt

## Pre-build Copy Files
    1. Copy pyadomd.txt content into pyadomd.py in the venv where pyinstaller will run
    copy pyadomd.txt venv\Lib\site-packages\pyadomd\pyadomd.py
    2. Copy custom.pth to venv site-packages
    copy custom.pth venv\Lib\site-packages\custom.pth

## Compile App
This app was built and tested with python 3.11.7
To create the .exe, normally you could run pyinstaller PowerBIRegressionTesterApp.py --onefile
However, since this app requires a copy of the ADOMD.dll a spec file is needed

    1. Create single EXE (recommended)
        To create a single executable:
        pyinstaller PowerBIRegressionTesterApp.spec --clean --noconfirm

    2. Create EXE with supporting folder #######
        To create an executable with the supporting files in a folder:
        pyinstaller PowerBIRegressionTesterApp-Folder.spec --clean --noconfirm

## Run App Manually
    1. To start the app manually, open the PowerBIRegressionTesterApp.py file and run from VS Code or run the compiled exe file