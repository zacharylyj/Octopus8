# Zachary Leong Octopus8 Mapper 5/9/2023
import pandas as pd
import json
import random


def apply_mappings(data_file_path, mappings):
    # Read the Excel sheet into a DataFrame
    df = pd.read_excel(data_file_path, engine='openpyxl')

    # Apply the mappings and create new columns
    for col_name, mapping in mappings.items():
        new_col_name = "Map " + col_name
        df[new_col_name] = df[col_name].map(mapping)
        print(f"{col_name}: Mapped")

    map_data_file_path = f"mapped_{data_file_path}"

    # Save the updated DataFrame back to the Excel file
    df.to_excel(map_data_file_path, index=False, engine='openpyxl')


if __name__ == "__main__":
    # Read the configuration from the JSON file
    with open('config.json', 'r') as config_file:
        config = json.load(config_file)

    # Extract values from the configuration
    excel_file_path = config["excel_file_path"]
    mappings = config["mappings"]

    # Useless Loading for fun :)
    print(f"Applying Map")
    loading = 0
    while loading < 100:
        print(f"{loading}%")
        loading += random.randint(0, 10)
    print(f"100%")

    # Call the function to apply the mappings and update the Excel file
    apply_mappings(excel_file_path, mappings)

    print(f"Mappings completed. Updated Excel file saved at {excel_file_path}")
