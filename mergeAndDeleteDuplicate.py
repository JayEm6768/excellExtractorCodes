import pandas as pd

# List of Excel files to merge (add your file names here)
excel_files = ['file1.xlsx', 'file2.xlsx']

# Read and concatenate all files
df_list = [pd.read_excel(f) for f in excel_files]
merged_df = pd.concat(df_list, ignore_index=True)

# Remove duplicate rows (keep the first occurrence)
merged_df = merged_df.drop_duplicates()

# Save the merged, deduplicated data to a new Excel file
merged_df.to_excel('merged_no_duplicates.xlsx', index=False)

print("✅ Files merged and duplicates removed. Output: 'merged_no_duplicates.xlsx'")
