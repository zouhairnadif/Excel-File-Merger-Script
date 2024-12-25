README: Excel File Merger Script
Overview

This Python script is designed to automate the process of merging multiple Excel files (both .xlsx and .xls formats) stored in separate month-specific folders into a single merged Excel file for each month. It reads all the Excel files in a given directory structure, processes each file, and combines them into a single file for each month. The script handles errors, such as file format issues, and logs them for further review.
Requirements

Before running this script, ensure the following Python packages are installed:

    pandas: To handle data manipulation and merging.
    openpyxl: Required for reading .xlsx files.
    xlrd: Required for reading .xls files.

You can install the required dependencies using pip:

pip install pandas openpyxl xlrd

Script Details
Input Folder Structure

    The script expects an input folder (defined as root_folder) containing subfolders named after the months (e.g., January, February, etc.).
    Each month folder should contain the Excel files to be merged. These files can be in .xlsx or .xls formats.

Example of folder structure:

root_folder
│
├── January
│   ├── file1.xlsx
│   ├── file2.xlsx
│   └── file3.xls
│
├── February
│   ├── file4.xlsx
│   └── file5.xls
└── March
    ├── file6.xlsx
    └── file7.xlsx

How the Script Works

    Root Folder and Month Folders:
        The script starts by defining the root folder, which contains subfolders for each month. Each subfolder represents a month's data and contains Excel files.

    Reading Excel Files:
        The script loops through each month folder and identifies Excel files (.xlsx and .xls).
        For each Excel file, the script attempts to read the content of the first sheet using the appropriate engine (openpyxl for .xlsx and xlrd for .xls).

    Merging Data:
        For each month folder, all valid Excel files are read and stored as DataFrames.
        The DataFrames are combined into a single DataFrame for the month.

    Error Handling:
        If an error occurs while reading a file (e.g., due to format issues), the script will log the error and continue processing the remaining files.

    Output File:
        After merging the data for a month, the script generates a single Excel file (<Month>_merged.xlsx) for that month and saves it in the root folder.
        The script appends the file name (without extension) as a new column in the DataFrame to track the source of each row.

Output

    The script generates a merged Excel file for each month folder.
    Each merged file contains data from all valid Excel files in that month's folder.

Example output file:

January_merged.xlsx
February_merged.xlsx
March_merged.xlsx

Each output file will have an additional column (Source_File) that indicates the name of the source file for each row.
Error Handling

    If the script encounters a ValueError while reading an Excel file (e.g., invalid or corrupted file format), it logs the error and skips the file.
    Other unexpected errors are also logged, and the script proceeds with the next file.

Customization

    Folder Path: Change the root_folder variable to specify the location of your root folder.
    File Types: The script currently handles .xlsx and .xls files. You can modify the file format checks if you need to handle other file types.

Example Usage

    Save the script as excel_merger.py.
    Adjust the root_folder variable to point to your folder containing the month folders.
    Run the script using Python:

python excel_merger.py

The script will process each month's folder, merge the Excel files, and save the merged files to the root folder.
License

This script is released under the MIT License. See the LICENSE file for more information.

Feel free to modify and adapt the script for your specific use case!