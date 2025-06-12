from PowerBIRegressionTester import PowerBIRegressionTester
import os
import sys
from tabulate import tabulate

working_directory = os.getcwd()
project_name = "Sales & Returns Sample"

tester = PowerBIRegressionTester.for_compare_only(project_name, working_directory)
df = None

user_value = input("Press enter to create a baseline, otherwise enter an instance name: ")

if tester.instance_exists(user_value):
    df = tester.compare(user_value.strip())
else:
    print("The instance does not exist.")
    sys.exit()

df = df[['visualId', 'RowCount', 'pageName']]
if df is not None:
    print(tabulate(df, headers='keys', tablefmt='psql', showindex=False))