import os
import glob
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import time

# Specify the path to the folder with PDF and Excel files
# folder_path = r'Licenses\Sasha'
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

# Open the first found Excel file using pandas
excel_file = excel_files[0]
df = pd.read_excel(excel_file)

# Select the GRZ column and convert it to a list
grz_values = df['GRZ'].tolist()

# Compare the GRZ values with the PDF file names
missing_values = [value for value in grz_values if value not in pdf_names]

if missing_values:
    print("GRZ values that are missing from the PDF file names:\n")
    for value in missing_values:
        print(value)
else:
    print("All GRZ values are present in the PDF file names. Everything is OK.")

time.sleep(300)