# GRZ_pdf_check_for_taxi_registry
## Description
This script helps verify the presence of Vehicle Registration Numbers (GRZ) from an Excel file against corresponding PDF filenames, which are expected to match these GRZ values. The primary goal of this script is to ensure that vehicles listed in the insurance company’s registry are accurately identified as either taxi or passenger cars. Since taxis have higher insurance rates than regular passenger cars, there’s a risk of fraud, where some taxi drivers might try to pass their vehicles off as regular cars to benefit from lower premiums.

To mitigate this risk, the script compares the GRZ values (vehicle registration numbers) listed in an Excel sheet with the filenames of downloaded PDF files from the taxi registry. It checks for any GRZ values missing from the PDFs and helps ensure that the correct vehicles are verified. If a taxi is identified under a regular car’s registration number, it allows insurance companies to take appropriate action by terminating the contract.

This automation helps prevent the loss of important GRZ data and provides an efficient way to cross-check the vehicle data in the insurance process, thereby supporting fraud prevention efforts.

## Functional Description
The program performs the following steps:
1. Retrieves the list of GRZ values from the Excel file.
2. Checks if each GRZ value is present in the corresponding PDF filenames.
3. Identifies missing GRZ values from the PDFs.
4. Outputs the results with missing GRZ values for further action.

## How It Works
1. The script loads the Excel file and extracts the GRZ values.
2. It retrieves the list of PDF files in the designated folder.
3. It compares the GRZ values with the names of the PDF files (which should match the GRZ values).
4. If any GRZ values are missing from the PDFs, they are displayed for further action.

## Input Structure
To run the program, the following parameters need to be provided:
1. Excel file containing the GRZ values.
2. Folder containing the PDF files named after GRZ values.

## Technical Requirements
To run the program, the following are required:
1. Python 3.x
2. Installed libraries: pandas, openpyxl, glob
3. The program assumes that the Excel file contains a column titled "GRZ" with vehicle registration numbers.

## Usage
1. Place the Excel file and PDF files in the designated folder.
2. Ensure the Excel file contains a column named "GRZ".
3. Run the script. It will check for missing GRZ values in the PDF filenames and display them if any are found.

## Example Output
Missing GRZ values:
- A list of GRZ values from the Excel file that are not found in the PDF filenames.

## Conclusion
This tool helps verify GRZ values from an insurance registry against PDF filenames to prevent fraud and ensure correct vehicle classification. It automates the process of matching vehicle registrations with the taxi registry, supporting more accurate data validation and fraud detection.
