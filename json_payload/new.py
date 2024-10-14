import os
import pandas as pd
from openpyxl import load_workbook

def excel_to_json(file_name, sheet_name):
    if os.path.exists(file_name):
        try:
            # Load the Excel file into a pandas DataFrame
            xl = pd.ExcelFile(file_name)

            if sheet_name in xl.sheet_names:
                # Read the sheet into a DataFrame
                df = xl.parse(sheet_name)

                # Assuming that the first two rows contain the table name and the actual data starts after that
                table_name = df.iloc[0, 0]  # First cell of the first row is the table name
                data_df = df.iloc[2:]  # Data starts from the 3rd row

                # Convert DataFrame to dictionary (records as list of dictionaries)
                data_dict = data_df.to_dict(orient='records')

                # Prepare the output as a dictionary
                result = {
                    "table_name": table_name,
                    "data": data_dict
                }

                # Convert dictionary to JSON format
                return result
            else:
                print(f"Sheet '{sheet_name}' not found in the file.")
                return None

        except Exception as e:
            print(f"Error reading the Excel file: {e}")
            return None
    else:
        print(f"File '{file_name}' does not exist.")
        return None

# Example usage
file_name = 'accountvault_data.xlsx'
sheet_name = input("Enter sheet name: ")

# Convert the sheet back to JSON
output = excel_to_json(file_name, sheet_name)

if output:
    print(output)
