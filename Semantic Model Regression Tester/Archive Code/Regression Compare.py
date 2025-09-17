from PowerBIRegressionTester import PowerBIRegressionTester
import sys
from tabulate import tabulate

project_folder = "Sales & Returns Sample"

tester = PowerBIRegressionTester.for_compare_only(project_folder)
df = None

instance_name = "T24"

if instance_name is None or instance_name == "":
    instance_name = input("Press enter to create a baseline, otherwise enter an instance name: ").strip()

if tester.instance_exists(instance_name):
    df = tester.compare(instance_name)
else:
    print("The instance does not exist.")
    sys.exit()

df = df[['pageName', 'visualId', 'RowCount']]
if df is not None:
    print(tabulate(df, headers='keys', tablefmt='psql', showindex=False))