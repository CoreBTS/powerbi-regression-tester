import json
import pandas as pd

json_path = "Sales & Returns Sample v201912.json"

with open(json_path, "r", encoding="utf-8-sig") as f:
    data = json.load(f)

# Load the 'events' list into a DataFrame
df = pd.DataFrame(data.get("events", []))


pd.set_option('display.max_columns', None)  # Show all columns
pd.set_option('display.max_rows', 20)    
print(df.head())

