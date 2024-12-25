import os
import pandas as pd

# Define the root folder where the monthly folders are stored
root_folder = r'C:\Users\ADMINTENOR\Desktop\excel merge files\OUT'

# List all the month folders (assuming they are named like January, February, etc.)
month_folders = os.listdir(root_folder)

# Loop over each folder (month)
for month in month_folders:
    month_folder_path = os.path.join(root_folder, month)

    # Check if it is a directory and not a file
    if os.path.isdir(month_folder_path):
        
        # Initialize a list to collect DataFrames for merging
        all_data = []

        # Loop over each Excel file in the month folder
        for file_name in os.listdir(month_folder_path):
            if file_name.endswith(('.xlsx', '.xls')):  # Check for both .xlsx and .xls files
                file_path = os.path.join(month_folder_path, file_name)

                # Determine the appropriate engine based on the file extension
                if file_name.endswith('.xlsx'):
                    engine = 'openpyxl'
                else:  # for .xls files
                    engine = 'xlrd'

                try:
                    # Attempt to read the Excel file with the correct engine
                    df = pd.read_excel(file_path, sheet_name=None, engine=engine)  # Read all sheets
                    print(f"Successfully read {file_name}")
                except ValueError as e:
                    # Handle the error when the file format cannot be determined
                    print(f"ValueError: Error reading {file_name}: {e}")
                    continue  # Skip this file and move to the next
                except Exception as e:
                    # Catch other potential exceptions
                    print(f"Unexpected error with file {file_name}: {e}")
                    continue  # Skip this file and move to the next

                # Assuming each file has only one sheet. If more, you can adjust the script.
                sheet_name = list(df.keys())[0]
                df = df[sheet_name]

                # Add a new column with the filename (without extension)
                df['Source_File'] = os.path.splitext(file_name)[0]

                # Append the DataFrame to the list
                all_data.append(df)

        # Concatenate all the DataFrames into one
        if all_data:  # Check if the list is not empty
            combined_df = pd.concat(all_data, ignore_index=True)

            # Define the output path for the merged Excel file
            output_file = os.path.join(root_folder, f"{month}_merged.xlsx")

            # Write the combined DataFrame to an Excel file (one sheet)
            combined_df.to_excel(output_file, index=False)

            print(f'Merged file for {month} saved to {output_file}')
        else:
            print(f"No valid Excel files to merge for {month}.")
