#! python3
# Calcium Imaging Analyzer v.1.73
# By Isaac Jardin

import pprint
import pyinputplus as pyip
from caiman_main_modules.caiman_analysis_module import complete_analysis_module
from caiman_main_modules.caiman_formating_module import complete_formating_module
from caiman_text import caiman_messages as ca_mess


print(ca_mess.welcome_message_v_1_73)


while True:
    menu_choice = pyip.inputMenu(choices=["Calcium mobilization analysis.",
                                          "File formating.",
                                          "Open Readme file",
                                          "Close the program."],
                                 prompt="What do you want to do?\n", numbered=True)
    if menu_choice == "Calcium mobilization analysis.":
        complete_analysis_module()

    elif menu_choice == "File formating.":
        complete_formating_module()

    elif menu_choice == "Open Readme file":
        with open("README.md", "r", encoding="UTF-8") as f:
            print(f.read())

    else:
        print("Thank you for using the program. Have a nice day.\n")
        break

# # TODO: Comentar cómo funciona cada cosa.
# # TODO: Hacer un file.bat para lanzarlo desde la consola.
