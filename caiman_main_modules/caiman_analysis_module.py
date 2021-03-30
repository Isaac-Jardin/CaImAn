#! python3
# Calcium Imaging Analyzer v.1.37
# By Isaac Jardin


import os
import pyinputplus as pyip
import time
from caiman_text import caiman_messages as ca_mess
from caiman_functions.caiman_analysis.caiman_ca_entry_analysis import calcium_entry_analysis, calcium_entry_multi_analysis
from caiman_functions.caiman_analysis.caiman_ca_oscillations_confocal import confocal_ca_oscillation_analysis
from caiman_functions.caiman_analysis.caiman_ca_oscillations_imaging import imaging_ca_oscillation_analysis, imaging_ca_oscillation_multi_analysis
from caiman_functions.caiman_analysis.caiman_soce_analysis import soce_analysis, soce_multi_analysis
from caiman_functions.caiman_utilities.csv_2_xlsx_converter import imageJ_reanalysis_multiple_files, imageJ_reanalysis_single_file
from pathlib import Path


def complete_analysis_module():
    while True:

        analysis_study = pyip.inputMenu(choices=["SOCE analysis (single file).",
                                                 "Calcium entry analysis (single file).",
                                                 "Imaging calcium oscillations analysis (single file).",
                                                 "Confocal calcium oscillations analysis (single file).",
                                                 "SOCE analysis (multiple files).",
                                                 "Calcium entry analysis (multiple files).",
                                                 "Imaging calcium oscillations analysis (multiple files).",
                                                 "Go back to the main menu.",
                                                 ], prompt="What analysis do you want to execute?: \n", numbered=True)

        if analysis_study == "SOCE analysis (single file).":
            route_input = pyip.inputFilepath(
                prompt="Please enter your file path: ")
            route = Path(route_input)
            os.chdir(route.parent)

            adquisiton_time = pyip.inputFloat(
                prompt="Introduce your acquisition time or leave it empty to set the default (2 seconds): ", blank=True)
            if adquisiton_time == "":
                adquisiton_time = 2
            integral_time = pyip.inputInt(
                prompt=f"""Set the number of frames(values) taken to calculate the AUC.
By default, 75 frames, which corresponds to 1 frame/2 seconds for 150 seconds.
Leave it empty to set the default value: """, blank=True)
            if integral_time == "":
                integral_time = 75
            pre_stimuli = pyip.inputInt(
                prompt="Introduce the ROW where the ratio basal pre-stimuli will be calculated: ", min=15)
            pre_calcium = pyip.inputInt(
                prompt="Introduce the ROW where the ratio basal pre-calcium will be calculated: ", min=15)

            slope_time_release = pyip.inputInt(
                prompt="""Set the number of frames(values) taken to calculate the calcium release trendline.
By default, 15 frames, which corresponds to 1 frame/2 seconds for 30 seconds.
Leave it empty to set the default value: """, blank=True)
            if slope_time_release == "":
                slope_time_release = 15

            slope_time_entry = pyip.inputInt(
                prompt="""Set the number of frames(values) taken to calculate the calcium entry trendline.
By default, 15 frames, which corresponds to 1 frame/2 seconds for 30 seconds.
Leave it empty to set the default value: """, blank=True)
            if slope_time_entry == "":
                slope_time_entry = 15

            start_time = time.time()
            soce_analysis(
                adquisiton_time, integral_time, pre_calcium, pre_stimuli, route, slope_time_entry, slope_time_release)
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
            integral_time = pyip.inputInt(
                prompt=f"""Set the number of frames(values) taken to calculate the AUC.
By default, 75 frames, which corresponds to 1 frame/2 seconds for 150 seconds.
Leave it empty to set the default value: """, blank=True)
            if integral_time == "":
                integral_time = 75
            pre_stimuli = pyip.inputInt(
                prompt="Introduce the ROW where the ratio basal pre-stimuli will be calculated: ", min=15)
            slope_time = pyip.inputInt(
                prompt="""Set the number of frames(values) taken to calculate the calcium entry trendline.
By default, 15 frames, which corresponds to 1 frame/2 seconds for 30 seconds.
Leave it empty to set the default value: """, blank=True)
            if slope_time == "":
                slope_time = 15
            start_time = time.time()
            calcium_entry_analysis(
                adquisiton_time, integral_time, pre_stimuli, route, slope_time)
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
            peak_amplitude = pyip.inputFloat(
                prompt="Introduce your sought peak amplitude (prominence) or leave it empty to set the default (0.02): ", blank=True)
            if peak_amplitude == "":
                peak_amplitude = 0.02

            peak_longitude = pyip.inputFloat(
                prompt="Introduce your sought peak longitude (width) or leave it empty to set the default (1): ", blank=True)
            if peak_longitude == "":
                peak_longitude = 1

            time_initial_linregress = pyip.inputInt(
                prompt=ca_mess.time_initial_linregress_text, min=15)

            time_final_linregress = pyip.inputInt(
                prompt=ca_mess.time_final_linregress_text, min=15)

            start_time = time.time()
            imaging_ca_oscillation_analysis(adquisiton_time, peak_amplitude, peak_longitude,
                                            route, time_initial_linregress, time_final_linregress)
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
            #     prompt=ca_mess.time_initial_linregress_text, min=15)

            # time_final_linregress = pyip.inputInt(
            #     prompt=ca_mess.time_final_linregress_text, min=15)

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

                        integral_time = pyip.inputInt(
                            prompt=f"""Set the number of frames(values) taken to calculate the AUC.
By default, 75 frames, which corresponds to 1 frame/2 seconds for 150 seconds.
Leave it empty to set the default value: """, blank=True)
                        if integral_time == "":
                            integral_time = 75

                        pre_stimuli = pyip.inputInt(
                            prompt="Introduce the ROW where the ratio basal pre-stimuli will be calculated: ", min=15)
                        pre_calcium = pyip.inputInt(
                            prompt="Introduce the ROW where the ratio basal pre-calcium will be calculated: ", min=15)
                        start_time = time.time()

                        slope_time_release = pyip.inputInt(
                            prompt="""Set the number of frames(values) taken to calculate the calcium release trendline.
By default, 15 frames, which corresponds to 1 frame/2 seconds for 30 seconds.
Leave it empty to set the default value: """, blank=True)
                        if slope_time_release == "":
                            slope_time_release = 15

                        slope_time_entry = pyip.inputInt(
                            prompt="""Set the number of frames(values) taken to calculate the calcium entry trendline.
By default, 15 frames, which corresponds to 1 frame/2 seconds for 30 seconds.
Leave it empty to set the default value: """, blank=True)
                        if slope_time_entry == "":
                            slope_time_entry = 15

                        soce_multi_analysis(
                            adquisiton_time, file, folder, integral_time, pre_calcium, pre_stimuli, slope_time_entry, slope_time_release)

                        print(
                            f"Execution time: {round((time.time()-start_time), 2)} seconds.")
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

                        integral_time = pyip.inputInt(
                            prompt=f"""Set the number of frames(values) taken to calculate the AUC.
By default, 75 frames, which corresponds to 1 frame/2 seconds for 150 seconds.
Leave it empty to set the default value: """, blank=True)
                        if integral_time == "":
                            integral_time = 75

                        pre_stimuli = pyip.inputInt(
                            prompt="Introduce the ROW where the ratio basal pre-stimuli will be calculated: ", min=15)

                        slope_time = pyip.inputInt(
                            prompt="""Set the number of frames(values) taken to calculate the calcium entry trendline.
By default, 15 frames, which corresponds to 1 frame/2 seconds for 30 seconds.
Leave it empty to set the default value: """, blank=True)
                        if slope_time == "":
                            slope_time = 15

                        start_time = time.time()
                        calcium_entry_multi_analysis(
                            adquisiton_time, file, folder, integral_time, pre_stimuli, slope_time)
                        print(
                            f"Execution time: {round((time.time()-start_time), 2)} seconds.")
                        print("Done.\n")

        elif analysis_study == "Imaging calcium oscillations analysis (multiple files).":
            print(
                "We are sorry, but we are still implenting this feature. Thank for your comprehension. Have a nice day")
            print()
            #         route_input = pyip.inputFilepath(
            #             prompt="Please enter your folder path: ")
            #         route_folder = Path(route_input)
            #         for folder, subfolders, files in os.walk(route_folder):
            #             for file in files:
            #                 if ".xlsx" in file and not "_modified" in file:
            #                     route = os.path.join(folder, file)
            #                     print(f"Please set the following conditions for {file}.")
            #                     os.chdir(route_folder.parent)
            #                     adquisiton_time = pyip.inputFloat(
            #                         prompt="Introduce your acquisition time or leave it empty to set the default (2 seconds): ", blank=True)
            #                     if adquisiton_time == "":
            #                         adquisiton_time = 2
            #                     peak_amplitude = pyip.inputFloat(
            #                         prompt="Introduce your sought peak amplitude (prominence) or leave it empty to set the default (0.02): ", blank=True)
            #                     if peak_amplitude == "":
            #                         peak_amplitude = 0.02

            #                     peak_longitude = pyip.inputFloat(
            #                         prompt="Introduce your sought peak longitude (width) or leave it empty to set the default (1): ", blank=True)
            #                     if peak_longitude == "":
            #                         peak_longitude = 1

            #                     time_initial_linregress = pyip.inputInt(
            #                         prompt="""
            # In order to work correctly, this program needs to calculate the fura2 fluorescence initial curve fitting (regression analysis).
            # Please introduce the ROW number previous to the cell response: """, min=15)

            #                     time_final_linregress = pyip.inputInt(
            #                         prompt=f"""
            # In order to work correctly, this program needs to calculate the fura2 fluorescence final curve fitting (regression analysis).
            # Please introduce the ROW number corresponding to 30-60 seconds before the end of your experiment: """)
            #                     if time_final_linregress == "":
            #                         peak_longitude = int(time_initial_linregress)

            #                     start_time = time.time()
            #                     imaging_ca_oscillation_multi_analysis(adquisiton_time, file, folder, peak_amplitude, peak_longitude,
            #                                                  route_input, time_initial_linregress, time_final_linregress)
            #                     print(
            #                         f"Execution time: {round((time.time()-start_time), 2)} seconds.")
            #                     print("Done.\n")

        else:
            print("Thank you for using the program. Have a nice day.")
            print()
            break


# # TODO: Comentar cómo funciona cada cosa.
# # TODO: Averiguar como arreglar lo de los dataframes.
# # TODO: Hacer un conversor de los ficheros del confocal.
# # TODO: Hacer un file.bat para lanzarlo desde la consola.
# # TODO: Hacer un vídeo explicativo.
