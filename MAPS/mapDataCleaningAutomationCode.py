import pandas as pd
import os

# --- CONFIGURATION ---
input_filename = "GT DP,NAP Utilization Report 20250715.xlsx"   # change to your actual file
output_dir = "output_files"
os.makedirs(output_dir, exist_ok=True)

# Columns to extract
columns_to_extract = [
    'DPdeniro', 'S_SP', 'S_Total', 'Com Date',
    'DP/NAP LAT', 'DP/NAP LONG', 'BRGY_NAME',
    'CFS Cluster', 'Tech', 'Location Type'
]

# Allowed clusters
valid_clusters = ["DAVAO NORTH", "DAVAO SOUTH", "TAGUM 1", "TAGUM 2"]

# Barangay groups
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
    'Alejandra Navarro', 'Tagpore',
    'Tibungco', 'Communal', 'San Isidro', 'Acacia','Tigatto'
]

group_mapping = {
    "South": group1_brgy,
    "Central": group2_brgy,
    "North": group3_brgy
}

# --- STEP 1: Load and filter data ---
df = pd.read_excel(input_filename)
df = df[columns_to_extract]
df = df[df['CFS Cluster'].isin(valid_clusters)]

# --- STEP 2: Split by barangay groups and save ---
def save_in_chunks(data, base_name):
    max_rows = 2000
    part = 0
    while len(data) > 0:
        part_df = data.iloc[:max_rows]
        data = data.iloc[max_rows:]
        
        if part == 0:
            filename = os.path.join(output_dir, f"{base_name}.xlsx")
        else:
            filename = os.path.join(output_dir, f"{base_name}_extended{part}.xlsx")
        
        part_df.to_excel(filename, index=False)
        part += 1

for name, brgys in group_mapping.items():
    filtered = df[df['BRGY_NAME'].isin(brgys)]
    if not filtered.empty:
        save_in_chunks(filtered, name)

# --- STEP 3: Create Spare files ---
# Rule: delete rows with VDSL, ADSL, ADSL/VDSL
#       replace blank (" ") Tech with "GPON"
for name in group_mapping.keys():
    files = [f for f in os.listdir(output_dir) if f.startswith(name) and f.endswith(".xlsx")]
    for file in files:
        file_path = os.path.join(output_dir, file)
        data = pd.read_excel(file_path)

        # Replace blank with GPON
        data['Tech'] = data['Tech'].replace(" ", "GPON")

        # Remove VDSL/ADSL/ADSL-VDSL
        data = data[~data['Tech'].isin(["VDSL", "ADSL", "ADSL/VDSL"])]

        spare_name = file.replace(".xlsx", " Spare.xlsx")
        data.to_excel(os.path.join(output_dir, spare_name), index=False)

# --- STEP 4: Create DSL file ---
all_data = []
for name in group_mapping.keys():
    files = [f for f in os.listdir(output_dir) if f.startswith(name) and f.endswith(".xlsx") and "Spare" not in f]
    for file in files:
        file_path = os.path.join(output_dir, file)
        data = pd.read_excel(file_path)
        all_data.append(data)

combined = pd.concat(all_data, ignore_index=True)
dsl_data = combined[combined['Tech'].isin(["VDSL", "ADSL", "ADSL/VDSL"])]
dsl_data.to_excel(os.path.join(output_dir, "DSL.xlsx"), index=False)

# --- STEP 5: Add coordinates column to all files ---
for file in os.listdir(output_dir):
    if file.endswith(".xlsx"):
        file_path = os.path.join(output_dir, file)
        data = pd.read_excel(file_path)
        if 'DP/NAP LAT' in data.columns and 'DP/NAP LONG' in data.columns:
            data['coordinates'] = data['DP/NAP LAT'].astype(str) + ", " + data['DP/NAP LONG'].astype(str)
        data.to_excel(file_path, index=False)
