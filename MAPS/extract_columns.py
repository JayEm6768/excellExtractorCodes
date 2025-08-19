# pip install pandas openpyxl

import pandas as pd
import os

# Paths
files_folder = "Files"
input_filename = "DVN DP Util 20250715.xlsx" # Change this to your actual input file name
output_filename = "DVN DP Util 20250715_extracted.xlsx"

# Ensure the files folder exists
os.makedirs(files_folder, exist_ok=True)

# Full paths
input_path = os.path.join(files_folder, input_filename)
output_path = os.path.join(files_folder, output_filename)

# List of columns you want to extract
columns_to_extract = [
    'DPdeniro', 'S_SP', 'S_Total', 'Com Date','DP/NAP LAT', 'DP/NAP LONG', 'BRGY_NAME'
]

# Step 1: Read the Excel file
df = pd.read_excel(input_path)

# Step 2: Clean column names
df.columns = df.columns.str.strip()

# Step 3: Check for missing columns
missing_cols = [col for col in columns_to_extract if col not in df.columns]
if missing_cols:
    print(f"⚠️ Warning: These columns are missing: {missing_cols}")
    # Only keep existing columns
    columns_to_extract = [col for col in columns_to_extract if col in df.columns]

# Step 4: Extract available columns
filtered_df = df[columns_to_extract]

# Step 5: Save to files folder
filtered_df.to_excel(output_path, index=False)

print(f"✅ Columns extracted and saved to '{output_path}'")
