# This script divides an Excel file into three equal parts.
# To use this script:
# 1. Make sure you have pandas and openpyxl installed:
#    pip install pandas openpyxl
# 2. Place the Excel file you want to split in the same directory as this script.
# 3. Change the value of the 'input_filename' variable to the name of your Excel file.
# 4. Run the script. It will create three new Excel files with the suffix _part1, _part2, and _part3.

import pandas as pd
import os
import numpy as np

# --- Configuration ---
# IMPORTANT: Replace with the name of your Excel file
# make sure that the file that you want to split is outside the "Files" folder
input_filename = "this.xlsx" 
# --- End Configuration ---

# Check if the input file name is provided
if not input_filename:
    print("Error: Please specify the input Excel file name in the 'input_filename' variable.")
else:
    try:
        # Create a 'Files' directory if it doesn't exist
        output_folder = "Files"
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Read the Excel file
        print(f"Reading '{input_filename}'...")
        df = pd.read_excel(input_filename)
        print("File read successfully.")

        # Get the total number of rows
        total_rows = len(df)
        print(f"Total rows in the file: {total_rows}")

        # Split the DataFrame into three parts
        print("Splitting the data into three parts...")
        df_parts = np.array_split(df, 3)

        # Get the base name and extension of the input file
        base_name, extension = os.path.splitext(input_filename)

        # Save each part to a new Excel file inside the 'Files' folder
        for i, df_part in enumerate(df_parts):
            part_number = i + 1
            output_filename = f"{base_name}_part{part_number}{extension}"
            output_path = os.path.join(output_folder, output_filename)
            
            print(f"Saving part {part_number} to '{output_path}'...")
            df_part.to_excel(output_path, index=False)
            print(f"Part {part_number} saved successfully.")

        print("\nScript finished successfully. The files have been saved in the 'Files' directory.")

    except FileNotFoundError:
        print(f"Error: The file '{input_filename}' was not found. Please make sure the file exists and the name is correct.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
