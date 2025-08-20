import pandas as pd

# --- Configuration ---
input_files = [
   "Davao_North_Only_combined.xlsx",
    "Group2_Jade_Valley_Tigatto_Airport_combined.xlsx",
    "Group3_Cabantian_Mandug_Panacan_combined.xlsx"
]
output_suffix = "_cleaned.xlsx"
# ----------------------

# Define non-breaking space
nbsp = "\u00A0"

# Keywords to remove
keywords = ["VDSL", "ADSL", "ADSL/VDSL"]

for input_file in input_files:
    # Load the file
    df = pd.read_excel(input_file)

    # Replace "no value" and blanks with non-breaking space
    df = df.replace("no value", nbsp, regex=False)
    df = df.fillna(nbsp)

    # Remove keywords & replace "/" with space
    for col in df.columns:
        if df[col].dtype == "object":  # only process text columns
            for word in keywords:
                df[col] = df[col].str.replace(word, "", regex=False)

            # Replace "/" with a space
            df[col] = df[col].str.replace("/", " ", regex=False)

    # Strip spaces but ensure cells don’t end up empty
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df = df.replace("", nbsp)

    # Save cleaned file
    output_file = input_file.replace(".xlsx", output_suffix)
    df.to_excel(output_file, index=False)
    print(f"✅ Cleaned {input_file} → {output_file}")
