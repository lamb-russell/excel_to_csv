"""
Module: excel_to_csv.py

This module provides functionality to extract data from an Excel workbook and save each worksheet as a separate CSV file.

Usage:
    python excel_to_csv.py <excel_file> [--output_dir <output_directory>]

Arguments:
    excel_file (str): Path to the Excel file.
    output_dir (str, optional): Output directory for the CSV files. If not provided, a directory with the same name as
    the Excel file (without the extension) will be created in the current working directory.

Example:
    python excel_to_csv.py path/to/your/excel_file.xlsx --output_dir path/to/output/directory

This module uses the following libraries:
    - os: For file and directory operations.
    - argparse: For handling command-line arguments.
    - pandas: For reading Excel files and writing CSV files.

The module performs the following steps:
    1. Parses the command-line arguments to get the Excel file path and output directory.
    2. Creates the output directory if it doesn't exist.
    3. Loads the Excel workbook using pandas.
    4. Iterates over each worksheet in the workbook.
    5. Extracts the data from each worksheet and saves it as a separate CSV file in the output directory.
    6. Preserves special characters in the CSV files.

Note:
    - The CSV files will be named using the base name of the Excel file followed by an underscore and the sheet name.
    - The module assumes that the Excel file has a valid format and structure.

"""
import os
import argparse
import pandas as pd

# Create an argument parser
parser = argparse.ArgumentParser(description='Extract data from an Excel workbook and save as CSV files.')
parser.add_argument('excel_file', help='Path to the Excel file')
parser.add_argument('--output_dir',
                    help='Output directory for the CSV files (default: directory with similar name as the Excel file)')

# Parse the command-line arguments
args = parser.parse_args()

# Get the Excel file path from the command-line argument
excel_file = args.excel_file

# Get the base name of the Excel file (without the extension)
base_name = os.path.splitext(os.path.basename(excel_file))[0]

# Set the output directory
if args.output_dir:
    output_directory = args.output_dir
else:
    output_directory = base_name

# Create the output directory (if it doesn't exist)
os.makedirs(output_directory, exist_ok=True)

# Load the Excel workbook
excel_data = pd.read_excel(excel_file, sheet_name=None)

# Iterate over each worksheet in the workbook
for sheet_name, data in excel_data.items():
    # Generate the CSV file name
    csv_file = f"{base_name}_{sheet_name}.csv"
    csv_path = os.path.join(output_directory, csv_file)

    # Save the worksheet data as a CSV file
    data.to_csv(csv_path, index=False, encoding='utf-8-sig')

    print(f"CSV file saved: {csv_path}")

