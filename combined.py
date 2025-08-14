# pip install pandas openpyxl

import pandas as pd
import os

# --- Configuration ---
files_folder = "Files"
input_filename = "GT DP,NAP Utilization Report 20250715.xlsx"  # Change to your actual input file
# --- End Configuration ---

# Ensure the files folder exists
os.makedirs(files_folder, exist_ok=True)

# Full paths
input_path = os.path.join(files_folder, input_filename)

# Columns to extract (including CFS Cluster for filtering)
columns_to_extract = [
    'DPdeniro', 'S_SP', 'S_Total', 'Com Date',
    'DP/NAP LAT', 'DP/NAP LONG', 'BRGY_NAME', 'CFS Cluster'
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

# Step 4: Filter by CFS Cluster (case-insensitive)
if 'CFS Cluster' in df.columns:
    df['CFS Cluster'] = df['CFS Cluster'].astype(str).str.strip().str.upper()
    clusters = ["DAVAO NORTH", "DAVAO SOUTH", "TAGUM 1", "TAGUM 2"]
    df_filtered = df[df['CFS Cluster'].isin(clusters)]
else:
    print("⚠️ 'CFS Cluster' column not found — no filtering applied.")
    df_filtered = df

# Step 5: Extract available columns
filtered_df = df_filtered[columns_to_extract]

# Step 6: Define barangay groups
group1_brgy = [
    'Agdao', 'Bago Gallera', 'Baliok', 'Bangkas Heights', 'Barangay 1-A', 
    'Barangay 2-A', 'Barangay 3-A', 'Barangay 4-A', 'Barangay 5-A', 'Barangay 6-A', 
    'Barangay 7-A', 'Barangay 8-A', 'Barangay 9-A', 'Barangay 10-A', 'Barangay 11-B', 
    'Barangay 12-B', 'Barangay 13-B', 'Barangay 14-B', 'Barangay 15-B', 'Barangay 16-B', 
    'Barangay 17-B', 'Barangay 18-B', 'Barangay 19-B', 'Barangay 20-B', 'Barangay 21-C', 
    'Barangay 22-C', 'Barangay 23-C', 'Barangay 24-C', 'Barangay 26-C', 'Barangay 27-C', 
    'Barangay 28-C', 'Barangay 29-C', 'Barangay 30-C', 'Barangay 31-D', 'Barangay 32-D', 
    'Barangay 33-D', 'Barangay 34-D', 'Barangay 35-D', 'Barangay 36-D', 'Barangay 37-D', 
    'Barangay 38-D', 'Barangay 39-D', 'Barangay 40-D', 'Bucana', 'Centro', 
    'Gov. Vicente Duterte', 'Gov. Paciano Bangoy', 'Lapu-lapu', 'Leon Garcia Sr.', 
    'San Antonio', 'Tres De Mayo', 'Zone 1',
    'Matina Crossing', 'Kap. Tomas Monteverde Sr.'
]

group2_brgy = [
    'Rafael Castillo', 'Sasa', 'Vicente Hizon Sr.', 
    'Ubalde', 'Wilfredo Aquino', 'Ilang', 'Pampanga',
    'Buhangin', 'Alfonso Angliongto Sr.'
]

group3_brgy = [
    'Cabantian', 'Mandug', 'Panacan', 'Bunawan', 'Indangan', 
    'Alejandra Navarro', 'Tagpore', 'Tibungco',
    'Communal', 'San Isidro', 'Acacia', 'Tigatto'
]

# Step 7: Save grouped data
output_path1 = os.path.join(files_folder, "Group1_Urban_Core_Bucana.xlsx")
output_path2 = os.path.join(files_folder, "Group2_Jade_Valley_Tigatto_Airport.xlsx")
output_path3 = os.path.join(files_folder, "Group3_Cabantian_Mandug_Panacan.xlsx")

df_group1 = filtered_df[filtered_df['BRGY_NAME'].isin(group1_brgy)]
df_group2 = filtered_df[filtered_df['BRGY_NAME'].isin(group2_brgy)]
df_group3 = filtered_df[filtered_df['BRGY_NAME'].isin(group3_brgy)]

df_group1.to_excel(output_path1, index=False)
df_group2.to_excel(output_path2, index=False)
df_group3.to_excel(output_path3, index=False)

print(f"✅ Group 1 saved: {output_path1} (Rows: {len(df_group1)})")
print(f"✅ Group 2 saved: {output_path2} (Rows: {len(df_group2)})")
print(f"✅ Group 3 saved: {output_path3} (Rows: {len(df_group3)})")
print("\n🎯 Script finished. Files are in the 'Files' folder.")
