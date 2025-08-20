import os
import pandas as pd
from fastapi import FastAPI, UploadFile
from fastapi.responses import FileResponse
import shutil
import tempfile
import zipfile

app = FastAPI()

# --- Your barangay groups (same as your code) ---
group1_brgy = [...]
group2_brgy = [...]
group3_brgy = [...]

group_mapping = {
    "South": group1_brgy,
    "Central": group2_brgy,
    "North": group3_brgy
}

columns_to_extract = [
    'DPdeniro', 'S_SP', 'S_Total', 'Com Date',
    'DP/NAP LAT', 'DP/NAP LONG', 'BRGY_NAME',
    'CFS Cluster', 'Tech', 'Location Type'
]

valid_clusters = ["DAVAO NORTH", "DAVAO SOUTH", "TAGUM 1", "TAGUM 2"]

def save_in_chunks(data, base_name, output_dir):
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

@app.post("/process")
async def process_excel(file: UploadFile):
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, file.filename)
        output_dir = os.path.join(tmpdir, "output_files")
        os.makedirs(output_dir, exist_ok=True)

        # Save upload
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # --- Your existing logic ---
        df = pd.read_excel(input_path)
        df = df[columns_to_extract]
        df = df[df['CFS Cluster'].isin(valid_clusters)]

        for name, brgys in group_mapping.items():
            filtered = df[df['BRGY_NAME'].isin(brgys)]
            if not filtered.empty:
                save_in_chunks(filtered, name, output_dir)

        # Spare files
        for name in group_mapping.keys():
            files = [f for f in os.listdir(output_dir) if f.startswith(name) and f.endswith(".xlsx")]
            for f in files:
                path = os.path.join(output_dir, f)
                data = pd.read_excel(path)
                data['Tech'] = data['Tech'].replace(" ", "GPON")
                data = data[~data['Tech'].isin(["VDSL", "ADSL", "ADSL/VDSL"])]
                spare_name = f.replace(".xlsx", " Spare.xlsx")
                data.to_excel(os.path.join(output_dir, spare_name), index=False)

        # DSL
        all_data = []
        for name in group_mapping.keys():
            files = [f for f in os.listdir(output_dir) if f.startswith(name) and f.endswith(".xlsx") and "Spare" not in f]
            for f in files:
                path = os.path.join(output_dir, f)
                data = pd.read_excel(path)
                all_data.append(data)
        combined = pd.concat(all_data, ignore_index=True)
        dsl_data = combined[combined['Tech'].isin(["VDSL", "ADSL", "ADSL/VDSL"])]
        dsl_data.to_excel(os.path.join(output_dir, "DSL.xlsx"), index=False)

        # Add coordinates
        for f in os.listdir(output_dir):
            if f.endswith(".xlsx"):
                path = os.path.join(output_dir, f)
                data = pd.read_excel(path)
                if 'DP/NAP LAT' in data.columns and 'DP/NAP LONG' in data.columns:
                    data['coordinates'] = data['DP/NAP LAT'].astype(str) + ", " + data['DP/NAP LONG'].astype(str)
                data.to_excel(path, index=False)

        # Zip all files
        zip_path = os.path.join(tmpdir, "result.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for f in os.listdir(output_dir):
                zipf.write(os.path.join(output_dir, f), arcname=f)

        return FileResponse(zip_path, filename="processed_files.zip")
