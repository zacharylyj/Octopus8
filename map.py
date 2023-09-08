import pandas as pd
import json
import random
import msoffcrypto


def decrypt_excel(file_path, password, decrypted_file_path):
    with open(file_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        with open(decrypted_file_path, "wb") as df:
            office_file.decrypt(df)


def apply_mappings(data_file_path, mappings):
    # Prompt for password
    password = input(
        "Enter the password for the Excel file (or press Enter if the file is not password-protected): ")

    # Decrypt file if password is provided
    decrypted_file_path = "decrypted_" + data_file_path
    if password:
        decrypt_excel(data_file_path, password, decrypted_file_path)
        df = pd.read_excel(decrypted_file_path, engine='openpyxl')
    else:
        df = pd.read_excel(data_file_path, engine='openpyxl')

    # Apply the mappings and create new columns
    for col_name, mapping in mappings.items():
        new_col_name = "Map " + col_name
        df[new_col_name] = df[col_name].map(mapping)
        print(f"{col_name}: Mapped")

    map_data_file_path = f"mapped_{data_file_path}"

    # Save the updated DataFrame back to Excel
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
