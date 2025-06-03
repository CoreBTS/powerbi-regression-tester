import json
import pandas as pd

# Path to your JSON file
json_path = "Sales & Returns Sample v201912.json"

# Read the JSON file
with open(json_path, "r", encoding="utf-8-sig") as f:
    data = json.load(f)

# Extract events
events = data.get("events", [])

# Flatten metrics into top-level columns
def flatten_event(event):
    flat = event.copy()
    metrics = flat.pop("metrics", {})
    flat.update(metrics)
    return flat

flat_events = [flatten_event(e) for e in events]

# Create DataFrame
df = pd.DataFrame(flat_events)

# Show the first few rows
print(df.head())