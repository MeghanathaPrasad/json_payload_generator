import pandas as pd
import json
# from openpyxl import load_workbook
#
# # Load the workbook and the sheet
# workbook = load_workbook("output.xlsx")
# sheet = workbook["Sheet5"]
#
# # Find the first non-empty cell in the sheet
# first_word = None
# for row in sheet.iter_rows(values_only=True):
#     for cell in row:
#         if cell is not None:
#             first_word = str(cell)
#             break
#     if first_word:
#         break

# Load Excel data into a DataFrame, skipping the header row
df = pd.read_excel('output.xlsx', sheet_name='Sheet3')

# Convert DataFrame to a list of dictionaries
data_dict = df.to_dict(orient='records')

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
