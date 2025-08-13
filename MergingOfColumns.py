# pip install pandas openpyxl

import pandas as pd
import os

# Paths
files_folder = "Files"
input_filename = "DVN DP Util 20250715_extracted.xlsx"
output_filename = "DVN DP Util 20250715_Final.xlsx"

# Full paths
input_path = os.path.join(files_folder, input_filename)
output_path = os.path.join(files_folder, output_filename)

# Read the Excel file
df = pd.read_excel(input_path)

# Clean column names (remove extra spaces)
df.columns = df.columns.str.strip()

# Check that required columns exist
required_cols = ["DP/NAP LAT", "DP/NAP LONG"]
missing_cols = [col for col in required_cols if col not in df.columns]
if missing_cols:
    raise ValueError(f"Missing required columns: {missing_cols}")

# Combine LAT and LONG into one column, skip NaNs
df["Coordinates"] = df.apply(
    lambda row: f"{row['DP/NAP LAT']},{row['DP/NAP LONG']}"
    if pd.notnull(row['DP/NAP LAT']) and pd.notnull(row['DP/NAP LONG'])
    else "",
    axis=1
)

# Save the result
df.to_excel(output_path, index=False)

print(f"✅ Coordinates merged and saved to '{output_path}'")
