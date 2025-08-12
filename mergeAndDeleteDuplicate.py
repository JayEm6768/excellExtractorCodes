# pip install pandas openpyxl

import pandas as pd
import os

# Folder where the Excel files are stored
files_folder = "Files"

# Ensure the folder exists
os.makedirs(files_folder, exist_ok=True)

# List of Excel files to merge (located inside the files folder)
excel_files = [
    os.path.join(files_folder, "file1.xlsx"), # Replace with actual file names
    os.path.join(files_folder, "file2.xlsx")  # Replace with actual file names
]

# Read and concatenate all files
df_list = [pd.read_excel(f) for f in excel_files]
merged_df = pd.concat(df_list, ignore_index=True)

# Remove duplicate rows (keep the first occurrence)
merged_df = merged_df.drop_duplicates()

# Output file path
output_path = os.path.join(files_folder, "merged_no_duplicates.xlsx")

# Save the merged, deduplicated data
merged_df.to_excel(output_path, index=False)

print(f"✅ Files merged and duplicates removed. Output: '{output_path}'")
