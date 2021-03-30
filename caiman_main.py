#! python3
# Calcium Imaging Analyzer v.1.69
# By Isaac Jardin


# import os
import pyinputplus as pyip
# import time
from caiman_main_modules.caiman_analysis_module import complete_analysis_module
from caiman_main_modules.caiman_formating_module import complete_formating_module
from caiman_text import caiman_messages as ca_mess


print(ca_mess.welcome_message_v_1_69)


while True:
    menu_choice = pyip.inputMenu(choices=["Calcium mobilization analysis.", "File formating.",
                                          "Close the program."], prompt="What do you want to do?\n", numbered=True)
    if menu_choice == "Calcium mobilization analysis.":
        complete_analysis_module()

    elif menu_choice == "File formating.":
        complete_formating_module()

    else:
        print("Thank you for using the program. Have a nice day.")
        print()
        break

        #
        #


# # TODO: Comentar cómo funciona cada cosa.
# # TODO: Hacer un conversor de los ficheros del confocal.
# # TODO: Hacer un file.bat para lanzarlo desde la consola.
