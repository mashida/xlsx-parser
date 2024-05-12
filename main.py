import shutil
import sys

import pandas as pd
import os
import zipfile
import logging

logger = logging.getLogger(__name__)
LOGGING_LEVEL = logging.DEBUG
logger.setLevel(LOGGING_LEVEL)


def get_script_folder():
    # path of main .py or .exe when converted with pyinstaller
    if getattr(sys, 'frozen', False):
        script_path = os.path.dirname(sys.executable)
    else:
        script_path = os.path.dirname(
            os.path.abspath(sys.modules['__main__'].__file__)
        )
    return script_path


def transform_string(input_string: str | int):
    if isinstance(input_string, int):
        input_string = str(input_string)
    # Step 1: Remove the quotation marks
    cleaned_string = input_string.replace('«', '').replace('»', '')
    cleaned_string = cleaned_string.replace('"', '')

    # Step 2: Trim any unwanted characters (if necessary)
    # In this case, no additional trimming is needed, but this step is included for completeness

    return cleaned_string


def read_excel_file(file_path, sheet: str):
    """
    Reads an Excel file and returns a DataFrame.
    """
    return pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet)


def create_excel_files(data, output_dir):
    """
    Creates new Excel files for each unique client_id using pandas.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Group the data by 'Клиент' and create a DataFrame for each group
    clients = data['№ плател.'].drop_duplicates()

    # Iterate over each group (DataFrame)
    for client in clients:
        # Create a new Excel file for each group
        file_name = f"{transform_string(client)}.xlsx"
        file_path = os.path.join(output_dir, file_name)

        # Write the group data to the Excel file
        data.loc[data['№ плател.'] == client,
        [
            'Накладная',
            'Дата',
            'Город отправления',
            'Город получения',
            'Вес,кг',
        ]
        ].to_excel(file_path,
                   index=False)


def zip_excel_files(output_dir, zip_file_name):
    """
    Zips all the new Excel files in the output directory.
    """
    with zipfile.ZipFile(zip_file_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, output_dir))


def remove_folder_with_contents(folder_path):
    try:
        # Check if the folder exists
        if os.path.exists(folder_path):
            # Remove the folder and all its contents
            shutil.rmtree(folder_path)
            print(f"Folder '{folder_path}' and all its contents have been removed.")
        else:
            print(f"The folder '{folder_path}' does not exist.")
    except PermissionError:
        print(f"Permission denied for folder '{folder_path}'.")


def main():
    # Get the directory of the current script
    script_directory = get_script_folder()
    print(f"{script_directory=}")
    os.chdir(script_directory)

    # Construct the path to the input.xlsx file
    input_file_path = os.path.join(script_directory, 'input.xlsx')
    print(f"{input_file_path=}")
    output_dir = 'output'
    zip_file_name = 'output.zip'
    sheet_name = 'Накладные'

    data = read_excel_file(input_file_path, sheet_name)

    create_excel_files(data, output_dir)
    # The rest of the script remains the same for zipping the files
    zip_excel_files(output_dir, zip_file_name)
    remove_folder_with_contents(output_dir)
    print(f"Excel files zipped successfully: {zip_file_name}")


if __name__ == "__main__":
    main()
