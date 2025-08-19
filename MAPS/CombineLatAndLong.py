import pandas as pd

# List your Excel files here
files = [
    "Davao_North_Only.xlsx",
    "Group2_Jade_Valley_Tigatto_Airport.xlsx",
    "Group3_Cabantian_Mandug_Panacan.xlsx"
]

for file in files:
    # Load the Excel file
    df = pd.read_excel(file)

    # Combine latitude and longitude into one column
    df["DP/NAP COORDINATES"] = df["DP/NAP LAT"].astype(str) + "," + df["DP/NAP LONG"].astype(str)

    # Save back to a new file
    output_file = file.replace(".xlsx", "_combined.xlsx")
    df.to_excel(output_file, index=False)

    print(f"Done! Saved combined file as {output_file}")
