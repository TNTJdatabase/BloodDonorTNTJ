import pandas as pd
import json

# Read CSV
df = pd.read_csv(
    r"C:\MEDICAL WING APP\Medical wing donor register\Tamilnadu_locations_database.csv",
    encoding="utf-8"
)

# Remove extra spaces
df["AREA"] = df["AREA"].astype(str).str.strip()
df["DISTRICT"] = df["DISTRICT"].astype(str).str.strip()
df["STATE"] = df["STATE"].astype(str).str.strip()

# Create display string
df["location"] = (
    df["AREA"] + ", " +
    df["DISTRICT"] + ", " +
    df["STATE"]
)

# Convert to JSON
records = df[["location"]].to_dict(orient="records")

# Save JSON
output_path = r"C:\MEDICAL WING APP\Medical wing donor register\TamilNadu_location_database.json"

with open(output_path, "w", encoding="utf-8") as f:
    json.dump(records, f, ensure_ascii=False, indent=4)

print("JSON Created Successfully!")
print("Saved to:", output_path)