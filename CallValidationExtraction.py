

import pandas as pd

# File path
excel_file = 'NewValidationAug18.xlsx'

# List of columns you want to extract
columns_to_extract = [
    'workordernumber',
    'customername',
    'customercontact',
    'customeraddress',
    'appointmentdate',
    'team',
    'delayreason',
    'facilityname'
]

# Step 1: Read the Excel file without usecols
df = pd.read_excel(excel_file)

# Step 2: Clean column names (strip whitespace and lower case for matching)
df.columns = df.columns.str.strip()

# Step 3: Check for missing columns
missing_cols = [col for col in columns_to_extract if col not in df.columns]
if missing_cols:
    print(f"⚠️ Warning: These columns are missing: {missing_cols}")
    # Only keep columns that exist to avoid crashing
    columns_to_extract = [col for col in columns_to_extract if col in df.columns]

# Step 4: Extract available columns
filtered_df = df[columns_to_extract]

# Step 5: Save to new Excel file
filtered_df.to_excel('CallValidation.xlsx', index=False)

print("✅ Columns extracted and saved to 'CallValidation.xlsx'")
