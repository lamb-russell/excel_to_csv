# Excel to CSV Converter

This repository contains a Python script that converts an Excel workbook into separate CSV files for each worksheet.

## Features

- Extracts data from each worksheet in an Excel workbook
- Saves each worksheet as a separate CSV file
- Preserves special characters in the CSV files
- Provides a command-line interface for easy usage

## Requirements

- Python 3.x
- pandas library

## Installation

1. Clone the repository
2. Change into the project directory
3. Install the required dependencies

```pip install -r requirements.txt```

## Usage

To convert an Excel workbook to CSV files, run the following command:
python excel_to_csv.py <excel_file> [--output_dir <output_directory>]

- `<excel_file>`: Path to the Excel file you want to convert.
- `<output_directory>` (optional): Output directory for the CSV files. If not provided, a directory with the same name as the Excel file (without the extension) will be created in the current working directory.

Example:
python excel_to_csv.py path/to/your/excel_file.xlsx --output_dir path/to/output/directory

## Output

The script will create a directory (if not specified) with the same name as the Excel file (without the extension) in the current working directory or the specified output directory. Inside that directory, it will generate separate CSV files for each worksheet in the Excel workbook.

The CSV files will be named using the base name of the Excel file followed by an underscore and the sheet name.
