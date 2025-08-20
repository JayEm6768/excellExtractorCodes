import pandas as pd

# --- Configuration ---
input_files = [
    "Davao_North_Only_combined_cleaned.xlsx",
    "Group2_Jade_Valley_Tigatto_Airport_combined_cleaned.xlsx",
    "Group3_Cabantian_Mandug_Panacan_combined_cleaned.xlsx"
]
output_file = "combined.xlsx"
# ----------------------

# Read and combine all files
dfs = [pd.read_excel(f) for f in input_files]
combined_df = pd.concat(dfs, ignore_index=True)

# Save the combined file
combined_df.to_excel(output_file, index=False)
print(f"✅ Combined {len(input_files)} files → {output_file}")
