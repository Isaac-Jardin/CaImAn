#! python3
# Calcium Imaging Analyzer v.1.37
# By Isaac Jardin


import os
import pyinputplus as pyip
import shutil
import time
from pathlib import Path

from caiman_text import caiman_messages as ca_mess
from caiman_functions.caiman_analysis.caiman_ca_entry_analysis import calcium_entry_analysis, calcium_entry_multi_analysis
from caiman_functions.caiman_analysis.caiman_ca_oscillations_confocal import confocal_ca_oscillation_analysis
from caiman_functions.caiman_analysis.caiman_ca_oscillations_imaging import imaging_ca_oscillation_analysis, imaging_ca_oscillation_multi_analysis
from caiman_functions.caiman_analysis.caiman_soce_analysis import soce_analysis, soce_multi_analysis
from caiman_functions.caiman_utilities.csv_2_xlsx_converter import imageJ_reanalysis_multiple_files, imageJ_reanalysis_single_file


def complete_analysis_module():
    while True:

        analysis_study = pyip.inputMenu(choices=ca_mess.analysis_study_dict,
                                        prompt="What analysis do you want to execute?: \n", numbered=True)

        if analysis_study == "SOCE analysis (single file).":
            route_input = pyip.inputFilepath(
                prompt="Please enter your file path: ")
            route = Path(route_input)
            os.chdir(route.parent)

            adquisiton_time = pyip.inputFloat(
                prompt="Introduce your acquisition time or leave it empty to set the default (2 seconds): ", blank=True)
            if adquisiton_time == "":
                adquisiton_time = 2

            keyword = pyip.inputStr(
                prompt=ca_mess.keyword_message, blank=True)
            if keyword == "":
                keyword = "Ratio"

            integral_time = pyip.inputInt(
                prompt=ca_mess.integral_time_message, blank=True)
            if integral_time == "":
                integral_time = 75

            pre_stimuli = pyip.inputInt(
                prompt="Introduce the ROW where the ratio basal pre-stimuli will be calculated: ", min=10)

            pre_calcium = pyip.inputInt(
                prompt="Introduce the ROW where the ratio basal pre-calcium will be calculated: ", min=10)

            slope_time_release = pyip.inputInt(
                prompt=ca_mess.slope_time_release_message, blank=True)
            if slope_time_release == "":
                slope_time_release = 15

            slope_time_entry = pyip.inputInt(
                prompt=ca_mess.slope_time_entry_message, blank=True)
            if slope_time_entry == "":
                slope_time_entry = 15

            start_time = time.time()
            soce_analysis(
                adquisiton_time, integral_time, keyword, pre_calcium, pre_stimuli, route, slope_time_entry, slope_time_release)
            print(
                f"Execution time: {round((time.time()-start_time), 2)} seconds.")
            print("Done.\n")

        elif analysis_study == "Calcium entry analysis (single file).":
            route_input = pyip.inputFilepath(
                prompt="Please enter your file path: ")
            route = Path(route_input)
            os.chdir(route.parent)

            adquisiton_time = pyip.inputFloat(
                prompt="Introduce your acquisition time or leave it empty to set the default (2 seconds): ", blank=True)
            if adquisiton_time == "":
                adquisiton_time = 2

            keyword = pyip.inputStr(
                prompt=ca_mess.keyword_message, blank=True)
            if keyword == "":
                keyword = "Ratio"

            integral_time = pyip.inputInt(
                prompt=ca_mess.integral_time_message, blank=True)
            if integral_time == "":
                integral_time = 75

            pre_stimuli = pyip.inputInt(
                prompt="Introduce the ROW where the ratio basal pre-stimuli will be calculated: ", min=10)

            slope_time = pyip.inputInt(
                prompt=ca_mess.slope_time_entry_message, blank=True)
            if slope_time == "":
                slope_time = 15

            start_time = time.time()
            calcium_entry_analysis(
                adquisiton_time, integral_time, keyword, pre_stimuli, route, slope_time)
            print(
                f"Execution time: {round((time.time()-start_time), 2)} seconds.")
            print("Done.\n")

        elif analysis_study == "Imaging calcium oscillations analysis (single file).":
            route_input = pyip.inputFilepath(
                prompt="Please enter your file path: ")
            route = Path(route_input)
            os.chdir(route.parent)

            adquisiton_time = pyip.inputFloat(
                prompt="Introduce your acquisition time or leave it empty to set the default (2 seconds): ", blank=True)
            if adquisiton_time == "":
                adquisiton_time = 2

            keyword = pyip.inputStr(
                prompt=ca_mess.keyword_message, blank=True)
            if keyword == "":
                keyword = "Ratio"

            peak_amplitude = pyip.inputFloat(
                prompt="Introduce your sought peak amplitude (prominence) or leave it empty to set the default (0.02): ", blank=True)
            if peak_amplitude == "":
                peak_amplitude = 0.02

            peak_longitude = pyip.inputFloat(
                prompt="Introduce your sought peak longitude (width) or leave it empty to set the default (1): ", blank=True)
            if peak_longitude == "":
                peak_longitude = 1

            time_initial_linregress = pyip.inputInt(
                prompt=ca_mess.time_initial_linregress_text, min=10)

            time_final_linregress = pyip.inputInt(
                prompt=ca_mess.time_final_linregress_text, min=10)

            y_min_value = pyip.inputFloat(
                prompt="Introduce your y axis minimum value to plot your data or leave it empty to let the program to calculate it: ", blank=True)

            y_max_value = pyip.inputFloat(
                prompt="Introduce your y axis maximum value to plot your data or leave it empty to let the program to calculate it: ", blank=True)

            start_time = time.time()
            imaging_ca_oscillation_analysis(adquisiton_time, keyword, peak_amplitude,
                                            peak_longitude, route, time_initial_linregress, time_final_linregress, y_max_value, y_min_value)
            print(
                f"Execution time: {round((time.time()-start_time), 2)} seconds.")
            print("Done.\n")

        elif analysis_study == "Confocal calcium oscillations analysis (single file).":
            print(
                "We are sorry, but we are still implenting this feature. Thank for your comprehension. Have a nice day")
            print()
            # route_input = pyip.inputFilepath(
            #     prompt="Please enter your file path: ")
            # route = Path(route_input)
            # os.chdir(route.parent)
            # adquisiton_time = pyip.inputFloat(
            #     prompt="Introduce your acquisition time or leave it empty to set the default (2 seconds): ", blank=True)
            # if adquisiton_time == "":
            #     adquisiton_time = 2
            # peak_amplitude = pyip.inputFloat(
            #     prompt="Introduce your sought peak amplitude (prominence) or leave it empty to set the default (0.02): ", blank=True)
            # if peak_amplitude == "":
            #     peak_amplitude = 0.02

            # peak_longitude = pyip.inputFloat(
            #     prompt="Introduce your sought peak longitude (width) or leave it empty to set the default (1): ", blank=True)
            # if peak_longitude == "":
            #     peak_longitude = 1

            # time_initial_linregress = pyip.inputInt(
            #     prompt=ca_mess.time_initial_linregress_text, min=10)

            # time_final_linregress = pyip.inputInt(
            #     prompt=ca_mess.time_final_linregress_text, min=10)

            # start_time = time.time()
            # confocal_ca_oscillation_analysis(adquisiton_time, peak_amplitude, peak_longitude,
            #                                  route, time_initial_linregress, time_final_linregress)
            # print(
            #     f"Execution time: {round((time.time()-start_time), 2)} seconds.")
            # print("Done.\n")

        elif analysis_study == "SOCE analysis (multiple files).":
            route_input = pyip.inputFilepath(
                prompt="Please enter your folder path: ")
            route_folder = Path(route_input)
            for folder, subfolders, files in os.walk(route_folder):
                for file in files:
                    if ".xlsx" in file:
                        route = os.path.join(folder, file)
                        print(
                            f"Please set the following conditions for {file}.")
                        adquisiton_time = pyip.inputInt(
                            prompt="Introduce your acquisition time or leave it empty to set the default (2 seconds): ", blank=True)
                        if adquisiton_time == "":
                            adquisiton_time = 2

                        keyword = pyip.inputStr(
                            prompt=ca_mess.keyword_message, blank=True)
                        if keyword == "":
                            keyword = "Ratio"

                        integral_time = pyip.inputInt(
                            prompt=ca_mess.integral_time_message, blank=True)
                        if integral_time == "":
                            integral_time = 75

                        pre_stimuli = pyip.inputInt(
                            prompt="Introduce the ROW where the ratio basal pre-stimuli will be calculated: ", min=10)
                        pre_calcium = pyip.inputInt(
                            prompt="Introduce the ROW where the ratio basal pre-calcium will be calculated: ", min=10)
                        start_time = time.time()

                        slope_time_release = pyip.inputInt(
                            prompt=ca_mess.slope_time_release_message, blank=True)
                        if slope_time_release == "":
                            slope_time_release = 15

                        slope_time_entry = pyip.inputInt(
                            prompt=ca_mess.slope_time_entry_message, blank=True)
                        if slope_time_entry == "":
                            slope_time_entry = 15

                        soce_multi_analysis(
                            adquisiton_time, file, folder, integral_time, keyword, pre_calcium, pre_stimuli, slope_time_entry, slope_time_release)

                        print(
                            f"Execution time: {round((time.time()-start_time), 2)} seconds.")
                        print("Done.\n")

        elif analysis_study == "SOCE analysis (Z-Stack).":
            route_input = pyip.inputFilepath(
                prompt="Please enter your folder path: ")
            route_folder = Path(route_input)
            print(
                f"Please set the following conditions for your Z-Stack experiment.")
            adquisiton_time = pyip.inputInt(
                prompt="Introduce your acquisition time or leave it empty to set the default (2 seconds): ", blank=True)
            if adquisiton_time == "":
                adquisiton_time = 2

            keyword = pyip.inputStr(
                prompt=ca_mess.keyword_message, blank=True)
            if keyword == "":
                keyword = "Ratio"

            integral_time = pyip.inputInt(
                prompt=ca_mess.integral_time_message, blank=True)
            if integral_time == "":
                integral_time = 75

            pre_stimuli = pyip.inputInt(
                prompt="Introduce the ROW where the ratio basal pre-stimuli will be calculated: ", min=10)
            pre_calcium = pyip.inputInt(
                prompt="Introduce the ROW where the ratio basal pre-calcium will be calculated: ", min=10)

            slope_time_release = pyip.inputInt(
                prompt=ca_mess.slope_time_release_message, blank=True)
            if slope_time_release == "":
                slope_time_release = 15

            slope_time_entry = pyip.inputInt(
                prompt=ca_mess.slope_time_entry_message, blank=True)
            if slope_time_entry == "":
                slope_time_entry = 15
            start_time = time.time()
            for folder, subfolders, files in os.walk(route_folder):
                for file in files:
                    if ".xlsx" in file:
                        route = os.path.join(folder, file)
                        soce_multi_analysis(
                            adquisiton_time, file, folder, integral_time, keyword, pre_calcium, pre_stimuli, slope_time_entry, slope_time_release)

            folder_output = Path(f"{route_folder.parent}\\XLSXs_analyzed")
            num_files = 0
            for folder, subfolders, files in os.walk(route_folder):
                for file in files:
                    if "analyzed" in file:
                        num_files += 1
                        source = os.path.join(route_folder, file)
                        route_output = os.path.join(folder_output, file)
                        shutil.move(source, route_output)
            print(
                f"Execution time: {round((time.time()-start_time), 2)} seconds.")
            print(f"Your {num_files} files have been moved to {folder_output}")
            print("Done.\n")

        elif analysis_study == "Calcium entry analysis (multiple files).":
            route_input = pyip.inputFilepath(
                prompt="Please enter your folder path: ")
            route_folder = Path(route_input)
            for folder, subfolders, files in os.walk(route_folder):
                for file in files:
                    if ".xlsx" in file:
                        route = os.path.join(folder, file)
                        print(
                            f"Please set the following conditions for {file}.")
                        adquisiton_time = pyip.inputInt(
                            prompt="Introduce your acquisition time or leave it empty to set the default (2 seconds): ", blank=True)
                        if adquisiton_time == "":
                            adquisiton_time = 2

                        keyword = pyip.inputStr(
                            prompt=ca_mess.keyword_message, blank=True)
                        if keyword == "":
                            keyword = "Ratio"

                        integral_time = pyip.inputInt(
                            prompt=ca_mess.integral_time_message, blank=True)
                        if integral_time == "":
                            integral_time = 75

                        pre_stimuli = pyip.inputInt(
                            prompt="Introduce the ROW where the ratio basal pre-stimuli will be calculated: ", min=10)

                        slope_time = pyip.inputInt(
                            prompt=ca_mess.slope_time_entry_message, blank=True)
                        if slope_time == "":
                            slope_time = 15

                        start_time = time.time()
                        calcium_entry_multi_analysis(
                            adquisiton_time, file, folder, integral_time, keyword, pre_stimuli, slope_time)
                        print(
                            f"Execution time: {round((time.time()-start_time), 2)} seconds.")
                        print("Done.\n")

        elif analysis_study == "Calcium entry analysis (Z-Stack).":
            route_input = pyip.inputFilepath(
                prompt="Please enter your folder path: ")
            route_folder = Path(route_input)
            print(
                f"Please set the following conditions for your Z-Stack experiment.")
            adquisiton_time = pyip.inputInt(
                prompt="Introduce your acquisition time or leave it empty to set the default (2 seconds): ", blank=True)
            if adquisiton_time == "":
                adquisiton_time = 2

            keyword = pyip.inputStr(
                prompt=ca_mess.keyword_message, blank=True)
            if keyword == "":
                keyword = "Ratio"

            integral_time = pyip.inputInt(
                prompt=ca_mess.integral_time_message, blank=True)
            if integral_time == "":
                integral_time = 75

            pre_stimuli = pyip.inputInt(
                prompt="Introduce the ROW where the ratio basal pre-stimuli will be calculated: ", min=10)

            slope_time = pyip.inputInt(
                prompt=ca_mess.slope_time_entry_message, blank=True)
            if slope_time == "":
                slope_time = 15

            start_time = time.time()
            for folder, subfolders, files in os.walk(route_folder):
                for file in files:
                    if ".xlsx" in file:
                        route = os.path.join(folder, file)

                        calcium_entry_multi_analysis(
                            adquisiton_time, file, folder, integral_time, keyword, pre_stimuli, slope_time)

            folder_output = Path(f"{route_folder.parent}\\XLSXs_analyzed")
            num_files = 0
            for folder, subfolders, files in os.walk(route_folder):
                for file in files:
                    if "analyzed" in file:
                        num_files += 1
                        source = os.path.join(route_folder, file)
                        route_output = os.path.join(folder_output, file)
                        shutil.move(source, route_output)
            print(
                f"Execution time: {round((time.time()-start_time), 2)} seconds.")
            print(f"Your {num_files} files have been moved to {folder_output}")
            print("Done.\n")

        elif analysis_study == "Imaging calcium oscillations analysis (multiple files).":

            route_input = pyip.inputFilepath(
                prompt="Please enter your folder path: ")
            route_folder = Path(route_input)
            for folder, subfolders, files in os.walk(route_folder):
                for file in files:
                    if ".xlsx" in file and not "_modified" in file:
                        route = os.path.join(folder, file)
                        print(
                            f"Please set the following conditions for {file}.")
                        os.chdir(route_folder.parent)
                        adquisiton_time = pyip.inputFloat(
                            prompt="Introduce your acquisition time or leave it empty to set the default (2 seconds): ", blank=True)
                        if adquisiton_time == "":
                            adquisiton_time = 2

                        keyword = pyip.inputStr(
                            prompt=ca_mess.keyword_message, blank=True)
                        if keyword == "":
                            keyword = "Ratio"

                        peak_amplitude = pyip.inputFloat(
                            prompt="Introduce your sought peak amplitude (prominence) or leave it empty to set the default (0.02): ", blank=True)
                        if peak_amplitude == "":
                            peak_amplitude = 0.02

                        peak_longitude = pyip.inputFloat(
                            prompt="Introduce your sought peak longitude (width) or leave it empty to set the default (1): ", blank=True)
                        if peak_longitude == "":
                            peak_longitude = 1

                        time_initial_linregress = pyip.inputInt(
                            prompt=ca_mess.time_initial_linregress_text, min=10)

                        time_final_linregress = pyip.inputInt(
                            prompt=ca_mess.time_final_linregress_text)
                        if time_final_linregress == "":
                            peak_longitude = int(time_initial_linregress)

                        y_min_value = pyip.inputFloat(
                            prompt="Introduce your y axis minimum value to plot your data or leave it empty to let the program to calculate it: ", blank=True)

                        y_max_value = pyip.inputFloat(
                            prompt="Introduce your y axis maximum value to plot your data or leave it empty to let the program to calculate it: ", blank=True)

                        start_time = time.time()
                        imaging_ca_oscillation_multi_analysis(adquisiton_time, keyword, file, folder, peak_amplitude,
                                                              peak_longitude, route, time_initial_linregress, time_final_linregress, y_max_value, y_min_value)
                        print(
                            f"Execution time: {round((time.time()-start_time), 2)} seconds.")
                        print("Done.\n")

        else:
            print("Thank you for using the program. Have a nice day.")
            print()
            break


# # TODO: Comentar cómo funciona cada cosa.
# # TODO: Averiguar como arreglar lo de los dataframes.
# # TODO: Hacer un conversor de los ficheros del confocal.
# # TODO: Hacer un file.bat para lanzarlo desde la consola.
# # TODO: Hacer un vídeo explicativo.
