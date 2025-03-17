import os
import glob
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import time

# Specify the path to the folder with PDF and Excel files
folder_path = os.path.dirname(__file__)

# Get a list of all PDF files in the folder
pdf_files = glob.glob(os.path.join(folder_path, '*.pdf'))

# Create a list of PDF file names without extensions
pdf_names = [os.path.splitext(os.path.basename(pdf_file))[0] for pdf_file in pdf_files]

# Get a list of all Excel files in the folder
excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

# Check if at least one Excel file is found
if not excel_files:
    print("No Excel files found in the folder.")
    exit()

# Process each Excel file found
for excel_file in excel_files:
    # Open the Excel file using pandas
    df = pd.read_excel(excel_file)

    # Check if the 'GRZ' column exists in the Excel file
    if 'GRZ' not in df.columns:
        print(f"Warning: 'GRZ' column not found in {excel_file}. Skipping this file.")
        continue

    # Select the 'GRZ' column and convert it to a list, excluding empty values
    grz_values = df['GRZ'].dropna().tolist()

    # Compare the 'GRZ' values with the PDF file names
    missing_values = [value for value in grz_values if value not in pdf_names]

    if missing_values:
        # If there are missing values, print them
        print(f"GRZ values that are missing from the PDF file names in {excel_file}:\n")
        for value in missing_values:
            print(value)
    else:
        # If all GRZ values are found, print a success message
        print(f"All GRZ values are present in the PDF file names in {excel_file}. Everything is OK.")

# Wait for 5 minutes (300 seconds) before the script finishes
time.sleep(300)
