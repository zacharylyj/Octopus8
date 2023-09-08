import pandas as pd
import json
import os
import msoffcrypto


def decrypt_excel(file_path, password, decrypted_file_path):
    with open(file_path, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        with open(decrypted_file_path, "wb") as df:
            office_file.decrypt(df)


def apply_filler(data_file_path, filler):
    # Prompt for password
    password = input(
        "Enter the password for the Excel file (or press Enter if the file is not password-protected): ")

    # Decrypt file if password is provided
    decrypted_file_path = "decrypted_" + data_file_path
    if password:
        decrypt_excel(data_file_path, password, decrypted_file_path)
        df = pd.read_excel(decrypted_file_path, engine='openpyxl')
        os.remove(decrypted_file_path)  # Delete the decrypted file
    else:
        df = pd.read_excel(data_file_path, engine='openpyxl')

    # Fill missing values with unique IDs
    for col_name, default_name in filler.items():
        counter = 1
        for i in range(len(df)):
            if pd.isna(df.at[i, col_name]):
                df.at[i, col_name] = f"{default_name}{counter}"
                counter += 1
        print(f"{col_name}: Filled empty cells")

    map_data_file_path = f"filled_{data_file_path}"

    # Save the updated DataFrame back to Excel
    df.to_excel(map_data_file_path, index=False, engine='openpyxl')


if __name__ == "__main__":
    # Read the configuration from the JSON file
    with open('config.json', 'r') as config_file:
        config = json.load(config_file)

    # Extract values from the configuration
    excel_file_path = config["excel_file_path"]
    filler = config["filler"]

    # Call the function to apply the mappings and update the Excel file
    apply_filler(excel_file_path, filler)

    print(
        f"Mappings and filler completed. Updated Excel file saved at {excel_file_path}")
