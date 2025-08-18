import pandas as pd
import os

# --- Configuration ---
input_filename = "DVN DP Util 20250715_extracted.xlsx" 
# --- End Configuration ---

# Define the groups based on BRGY_NAME
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
    # Added
    'Matina Crossing',
    'Kap. Tomas Monteverde Sr.'
]

group2_brgy = [
     'Rafael Castillo', 'Sasa', 'Vicente Hizon Sr.', 
    'Ubalde', 'Wilfredo Aquino', 'Ilang', 'Pampanga',
    # Added
    'Buhangin',
    'Alfonso Angliongto Sr.'
]

group3_brgy = [
    'Cabantian', 'Mandug', 'Panacan', 'Bunawan', 'Indangan', 
    'Alejandra Navarro', 'Tagpore',
    # Added
    'Tibungco',
    'Communal',
    'San Isidro',
    'Acacia','Tigatto'
]

if not input_filename:
    print("Error: Please specify the input Excel file name in the 'input_filename' variable.")
else:
    try:
        output_folder = "Files"
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        print(f"Reading '{input_filename}'...")
        df = pd.read_excel(input_filename)
        print("File read successfully.")

        df_group1 = df[df['BRGY_NAME'].isin(group1_brgy)]
        output_path1 = os.path.join(output_folder, "Group1_Urban_Core_Bucana.xlsx")
        df_group1.to_excel(output_path1, index=False)
        print(f"Group 1 saved successfully. Rows: {len(df_group1)}")

        df_group2 = df[df['BRGY_NAME'].isin(group2_brgy)]
        output_path2 = os.path.join(output_folder, "Group2_Jade_Valley_Tigatto_Airport.xlsx")
        df_group2.to_excel(output_path2, index=False)
        print(f"Group 2 saved successfully. Rows: {len(df_group2)}")

        df_group3 = df[df['BRGY_NAME'].isin(group3_brgy)]
        output_path3 = os.path.join(output_folder, "Group3_Cabantian_Mandug_Panacan.xlsx")
        df_group3.to_excel(output_path3, index=False)
        print(f"Group 3 saved successfully. Rows: {len(df_group3)}")

        print("\nScript finished successfully. The files have been saved in the 'Files' directory.")

    except FileNotFoundError:
        print(f"Error: The file '{input_filename}' was not found. Please make sure the file exists and the name is correct.")
    except KeyError:
        print("Error: The 'BRGY_NAME' column was not found in the Excel file. Please check the column name.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
