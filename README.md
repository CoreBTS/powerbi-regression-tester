# Power BI Regression Tester

The Power BI Regression Tester application will allow you to test changes made to a model to see if the change made a difference. This can be useful when trying to optimze DAX, change settings such as Filter Value Behavior switch or changing a relationship.

The application requires have access to two models that can be compared and input a set of Power BI DAX query files either frmo Power BI Performance Analyzer or a DAX Studio Trace. 

These queries can then be run on both models to see if there are differences. If there are differences, the individual queries can be run to see what the differences are both at the row level and column level.


## Run App executable
    1. To start the app manually, open the PowerBIRegressionTesterApp\PowerBIRegressionTesterApp.exe file

## Setup App from the Project
    1. Clone the project
    2. git clone https://github.com/CoreBTS/powerbi-regression-tester.git
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

## Using the App

### Step 1: Create a Project
- Launch the app and create a new **Project** to group your regression test. A project represents a single testing effort (e.g., testing a DAX optimization or a relationship change).

### Step 2: Create a Baseline Connection
- Within your project, create a **Baseline** entry pointing to the **original/unchanged model**.
- Enter the connection properties for the model (e.g., a local Power BI Desktop instance or a published Power BI workspace).
- This is the "before" snapshot that all comparisons will be measured against.

### Step 3: Create an Instance Connection
- Create an **Instance** entry pointing to the **modified/changed model**.
- Enter the connection properties for the changed version of the model.
- This is the "after" snapshot to compare against the baseline.

### Step 4: Capture DAX Queries
Capture the queries you want to test using one of the following methods:

- **Power BI Performance Analyzer** — In Power BI Desktop, open the Performance Analyzer pane (View → Performance Analyzer), click **Start recording**, interact with the report visuals, then click **Export** to save the results as a JSON file.
- **DAX Studio Trace** — Open DAX Studio, connect to your model, start a trace with All Queries, interact with the report to generate queries, then export the trace results.

### Step 5: Import Queries
- In the app, use the **Import Queries** function for your project.
- Select the exported file from Power BI Performance Analyzer or DAX Studio.
- The app will parse the file and load the individual DAX queries into the project's **Query Files** folder.

### Step 6: Run Queries Against Baseline and Instance
- Run the imported queries against the **Baseline** model to capture the baseline results.
- Run the imported queries against the **Instance** model to capture the changed results.
- Both must complete successfully before a comparison can be performed.

### Step 7: Run the Comparison
- Once both the baseline and instance have results, click **Run Compare**.
- The app will compare the results query by query and report:
  - Queries with **no differences** (pass)
  - Queries with **differences**, including row-level and column-level details
- Review the differences to determine whether the model change had the intended effect or introduced regressions.

