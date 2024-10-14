import pandas as pd
import json

def read_excel_sheets(file_path):
    # Load the excel file and get sheet names
    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names

        # Display sheet names as choices
        print("Available sheets in the Excel file:")
        for idx, sheet in enumerate(sheet_names):
            print(f"{idx + 1}. {sheet}")

        # Take user input to select a specific sheet
        while True:
            try:
                choice = int(input("Enter the sheet number you want to read: "))
                if 1 <= choice <= len(sheet_names):
                    sheet_name = sheet_names[choice - 1]
                    break
                else:
                    print(f"Invalid input! Please enter a number between 1 and {len(sheet_names)}.")
            except ValueError:
                print("Please enter a valid number.")

        # Load and return the selected sheet
        print(f"Loading sheet: {sheet_name}")
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df

    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

# File path input
file_path = input("Enter the Excel file path: ")

# Call the function to read the selected sheet
dataframe = read_excel_sheets(file_path)

if dataframe is not None:
    print("\nLoaded data from selected sheet:")
    # Convert DataFrame to a list of dictionaries
    data_dict = dataframe.to_dict(orient='records')

    # Initialize the nested dictionary
    nested_dict = {}


    # Function to update the nested dictionary
    def update_nested_dict(keys, value, current_dict):
        for i, key in enumerate(keys):
            if i == len(keys) - 1:
                # Final key, set the value
                if isinstance(current_dict, dict):
                    current_dict[key] = value
                elif isinstance(current_dict, list):
                    current_dict.append(value)
            else:
                if key.isdigit():  # Handle array indices
                    key = int(key)
                    if len(current_dict) <= key:
                        current_dict.append({})
                    current_dict = current_dict[key]
                else:
                    if key not in current_dict:
                        current_dict[key] = [] if keys[i + 1].isdigit() else {}
                    current_dict = current_dict[key]


    # Process each row from the DataFrame
    for entry in data_dict:
        keys = entry['Key'].split('.')
        update_nested_dict(keys, entry['Value'], nested_dict)

    # Final result
    # result = {first_word: nested_dict}

    # Save the result to a JSON file
    with open('result.json', 'w') as file:
        json.dump(nested_dict, file, indent=4)

    print(nested_dict)
else:
    print("No data found")