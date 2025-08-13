# pip install pandas openpyxl

import pandas as pd
import os

# Path to your files folder and Excel file
files_folder = "Files"
input_filename = "DavaoNorthUtilExtract.xlsx"

# Full path
input_path = os.path.join(files_folder, input_filename)

# Read the Excel file
df = pd.read_excel(input_path)

# Clean column names
df.columns = df.columns.str.strip()

# Check if 'DPdeniro' column exists
if 'DPdeniro' in df.columns:
    unique_count = df['DPdeniro'].nunique(dropna=True)
    print(f"✅ There are {unique_count} unique DPdeniro values.")
else:
    print("⚠️ Column 'DPdeniro' not found in the file.")
