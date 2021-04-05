#! python3
# Calcium Imaging Analyzer v.1.37
# By Isaac Jardin


import os
import pyinputplus as pyip
import time
from caiman_text import caiman_messages as ca_mess
from caiman_functions.caiman_utilities.csv_2_xlsx_converter import imageJ_reanalysis_multiple_files, imageJ_reanalysis_single_file
from caiman_functions.caiman_utilities.filename_without_spaces import non_space_filename_single_file, non_space_filename_multiple_files
from pathlib import Path


def complete_formating_module():

    while True:

        analysis_study = pyip.inputMenu(choices=["_imageJ.csv to .xlsx file converter (single file).",
                                                 ".xlsx filename without spaces (single file).",
                                                 "_imageJ.csv to .xlsx files converter (multiple files).",
                                                 ".xlsx filenames without spaces (multiple file).",
                                                 "Go back to the main menu.",
                                                 ], prompt="What sort of file do you want to format?: \n", numbered=True)

        if analysis_study == "_imageJ.csv to .xlsx file converter (single file).":
            print(ca_mess.imageJ_csv_to_xlsx_message)
            route = pyip.inputFilepath(
                prompt="Please enter your file path: ")
            start_time = time.time()
            imageJ_reanalysis_single_file(route)
            print(
                f"Execution time: {round((time.time()-start_time), 2)} seconds.")
            print("Done.\n")

        elif analysis_study == ".xlsx filename without spaces (single file).":
            print(ca_mess.filename_without_spaces_message)
            route = pyip.inputFilepath(
                prompt="Please enter your file path: ")
            start_time = time.time()
            non_space_filename_single_file(route)
            print(
                f"Execution time: {round((time.time()-start_time), 2)} seconds.")
            print("Done.\n")

        elif analysis_study == "_imageJ.csv to .xlsx files converter (multiple files).":
            print(ca_mess.imageJ_csv_to_xlsx_message)
            route = pyip.inputFilepath(
                prompt="Please enter your folder path: ")
            start_time = time.time()
            imageJ_reanalysis_multiple_files(route)
            print(
                f"Execution time: {round((time.time()-start_time), 2)} seconds.")
            print("Done.\n")

        elif analysis_study == ".xlsx filenames without spaces (multiple file).":
            print(ca_mess.filename_without_spaces_message)
            route = pyip.inputFilepath(
                prompt="Please enter your folder path: ")
            start_time = time.time()
            non_space_filename_multiple_files(route)
            print(
                f"Execution time: {round((time.time()-start_time), 2)} seconds.")
            print("Done.\n")

        else:
            break


# # TODO: Comentar cómo funciona cada cosa.
# # TODO: Averiguar como arreglar lo de los dataframes.
# # TODO: Hacer un conversor de los ficheros del confocal.
# # TODO: Hacer un file.bat para lanzarlo desde la consola.
# # TODO: Hacer un vídeo explicativo.
