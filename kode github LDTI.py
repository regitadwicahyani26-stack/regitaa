import pandas as pd
import json

def convert_json_to_excel(json_file_path, excel_file_path):
    """
    Converts a JSON file containing multiple datasets into an Excel file,
    with each dataset placed on a separate sheet.
    
    Args:
        json_file_path (str): The path to the input JSON file.
        excel_file_path (str): The desired path for the output Excel file.
    """
    try:
        # Load the JSON data from the file
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Create a Pandas Excel writer object
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            # Iterate through each key-value pair in the JSON data
            for sheet_name, sheet_data in data.items():
                print(f"Converting data for sheet: '{sheet_name}'...")
                
                # Convert the list of dictionaries to a pandas DataFrame
                df = pd.DataFrame(sheet_data)
                
                # Write the DataFrame to a specific sheet in the Excel file
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\nSuccessfully converted '{json_file_path}' to '{excel_file_path}'. 🎉")
    
    except FileNotFoundError:
        print(f"Error: The file '{json_file_path}' was not found.")
    except json.JSONDecodeError:
        print(f"Error: The file '{json_file_path}' is not a valid JSON file.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# --- Example Usage ---
# Assuming your JSON file is named 'REGITA DWI CAHYANI_V3925051_TI-A_LDTI.json'
# and you want to save the Excel file as 'LDTI_Data.xlsx'
json_input_file = 'REGITA DWI CAHYANI_V3925051_TI-A_LDTI.json'
excel_output_file = 'LDTI_Data.xlsx'

convert_json_to_excel(json_input_file, excel_output_file)
