# GS1 results processor
## Overview

This script is designed to process multiple Excel files from GS1 portal. Those files are results
after importing files with EAN codes. This script extracts specific data and compiling this data into a new Excel file.

## How it works

1. The script scans a designated folder for Excel files.
2. For each Excel file found, it checks each row starting from a specified row number.
3. It evaluates a condition based on the contents of a specific column (CONDITIONS_COLUMN).
4. If the condition matches one of two predefined conditions (CONDITION_1 or CONDITION_2), data from specified
5. This extracted data is collected from all processed files.
6. Finally, the script creates a new Excel file and writes all the collected data into it.

## Usage

1. Set the folder_path variable to the directory containing the Excel files to be processed.
2. Set the output_file variable to the desired name for the output Excel file.
3. Run the script. It will process all Excel files in the specified directory, extract the required data, and save it in a new file named as per output_file.