import openpyxl
import os
import pandas as pd
from pathlib import Path


# non_space_filename_single_file takes a given Excel file (.xlsx) with spaces in its filename,
# and created a copy of it with a new filename where the " " have been replaced by "_"
def non_space_filename_single_file(route):
    file_path = Path(route)
    wb = openpyxl.load_workbook(file_path)
    splitted_file = file_path.stem
    non_space_filename = splitted_file.replace(" ", "_")
    xlsx_append = f"{non_space_filename}.xlsx"
    save_route = os.path.join(file_path.parent, xlsx_append)
    wb.save(save_route)
    print(f"'{file_path.name}' has been converted to '{xlsx_append}'.")


# non_space_filename_multiple_files search for Excel files (.xlsx) with spaces in their filename in a given path,
# and created copies of them with a new filename where the " " have been replaced by "_"
def non_space_filename_multiple_files(route):
    for filename in os.listdir(route):
        if " " in filename and ".xlsx" in filename:
            file_path = os.path.join(route, filename)
            wb = openpyxl.load_workbook(file_path)
            non_space_filename = filename.replace(" ", "_")
            non_space_filename_path = os.path.join(route, non_space_filename)
            print(non_space_filename_path)
            wb.save(non_space_filename_path)
            print(f"'{filename}' has been converted to '{non_space_filename}'.")
