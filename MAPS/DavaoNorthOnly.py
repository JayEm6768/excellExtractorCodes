# pip install pandas openpyxl

import pandas as pd
import os

# --- Configuration ---
files_folder = "Files"
input_filename = "Group1_Urban_Core_Bucana.xlsx"  # Change to your actual input file
# --- End Configuration ---

# Ensure the files folder exists
os.makedirs(files_folder, exist_ok=True)

# Full paths
input_path = os.path.join(files_folder, input_filename)

# Columns to extract (including CFS Cluster for filtering)
columns_to_extract = [
    'DPdeniro', 'S_SP', 'S_Total', 'Com Date',
    'DP/NAP LAT', 'DP/NAP LONG', 'BRGY_NAME', 'CFS Cluster', 'Tech', 'Location Type'
]

# Step 1: Read the Excel file
df = pd.read_excel(input_path)

# Step 2: Clean column names
df.columns = df.columns.str.strip()

# Step 3: Check for missing columns
missing_cols = [col for col in columns_to_extract if col not in df.columns]
if missing_cols:
    print(f"⚠️ Warning: These columns are missing: {missing_cols}")
    columns_to_extract = [col for col in columns_to_extract if col in df.columns]

# Step 4: Filter by only "DAVAO NORTH" (case-insensitive)
if 'CFS Cluster' in df.columns:
    df['CFS Cluster'] = df['CFS Cluster'].astype(str).str.strip().str.upper()
    df_filtered = df[df['CFS Cluster'] == "DAVAO NORTH"]
else:
    print("⚠️ 'CFS Cluster' column not found — no filtering applied.")
    df_filtered = df

# Step 5: Extract available columns
filtered_df = df_filtered[columns_to_extract]

# Step 6: Save the filtered data
output_path = os.path.join(files_folder, "Davao_North_Only.xlsx")
filtered_df.to_excel(output_path, index=False)

print(f"✅ Filtered data saved: {output_path} (Rows: {len(filtered_df)})")
