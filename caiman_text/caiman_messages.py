welcome_message_v_1_0 = """
    Welcome to Calcium Imaging Analyzer v.1.0
    By Isaac Jardin.

    By running this program you will be able to automatically analyze your fura2 experiments recorded with the Nikon Imaging Software (NIS).
    Before you begin, you need to consider few things that you must set within the program:

    1 - Are you analizing a single file or a whole set of experiments in a folder?. You will need to provide a path for your file/s.

    2 - Are you analizing Calcium entry (Medium with calcium -> Agonist) or
        classical SOCE experiments (Calcium-free medium [EGTA] -> Agonist -> Extracellular calcium)?.

    3 - The default adquisiton time is set to 2 seconds. If you wish, you would be able to manually modify it.

    4 - By default, this program calculates the agonist-evoked Ca2+ release and/or entry as the integral of the rise in fura-2 fluorescence ratio
        for 2½ min after the addition of the agonist or CaCl2.  
        To achieve that, the program will calculate the average ratio value from the 10 previous ratio values (basal) before stimulus or calcium
        addition. Therefore, you must previously check your excel file to get the ROW NUMBER where your stimulation/calcium adition starts.
        During the program execution:
            
            If you have chosen 'Calcium entry', you would need to provide ONE ROW NUMBER (Basal pre-stimuli).
            If you have chosen 'Classical SOCE', you would need to provide TWO ROW NUMBERS (Basal pre-stimuli and basal pre-calcium adition).
        
        The default integral time is set to 75, which corresponds to 150 seconds in 2½ min divided by 2 seconds (adquisition time).
        IF YOU HAVE MODIFIED THE ADQUISITION TIME, YOU SHOULD CALCULATE THE CORRECT INTEGRAL TIME AND INTRODUCE IT,
        OR YOUR RESULT WILL BE INCORRECT.

    5 - After your input, the program will calculate the agonist-evoked Ca2+ release and/or entry for each individual cell as well as
        the maximum peaks, slopes.
        Additionaly, The average values and standard errors for each parameters from the whole experiment will be determined.        
        Moreover, the program will generate a second worksheet, containing the same data than above but normalized to F/F0.
        It will also generates two scattercharts, one on each worksheet, including the traces from all the analyzed cells.

    6 - Finally, the program will create a copy of your file/files with the calculations and the tag '_analyzed' at the end,
        leaving your original file unmodified. 

    If you have doubts or detect a bug, please contact me at ijp@unex.es.
    Have fun.!
"""

welcome_message_v_1_37 = """
    Welcome to Calcium Imaging Analyzer v.1.37
    By Isaac Jardin.

    By running this program you will be able to automatically analyze your imaging experiments recorded with the Nikon Imaging Software (NIS).

    1 - Analysis:
            > Calcium entry experiments (Medium with calcium -> Agonist). Single and multiple files.
            > Classical SOCE experiments (Calcium-free medium [EGTA] -> Agonist -> Extracellular calcium). Single and multiple files.
            > Calcium oscilation experiments (Medium with calcium -> Agonist). Single files.
        
        Formating:
            > CSV files generated with the Fiji ImageJ macro "Reanálisis_Fura2_Nikon_FINAL_V_2.ijm" modified from Pedro Camello.

        Future upgrades:
            > Calcium oscilation experiments (Medium with calcium -> Agonist). Multiple files.
            > Formating Nikon confocal files.
            > Formating Metafluor files.
    
    2 - The default adquisiton time is set to 2 seconds. If you wish, you would be able to manually modify it.
        Calcium Imaging Analyzer calculates the agonist-evoked Ca2+ release and/or entry as the integral of the rise in fura-2 fluorescence ratio
        for 2½ min after the addition of the agonist or CaCl2.  
        To achieve that, the program will estimate the average ratio value from the 10 previous ratio values (basal) before stimulus or calcium
        addition. Therefore, you must previously check your excel file to get the ROW NUMBER where your stimulation/calcium adition starts to introduce them.
        Furthermore, the program will use those values as the initial position to calculate the slope of your traces. By default, it will take at least
        15 of the following values. You might increase this number manually to adjust it to your experiment.  
        After your input, the program will calculate the agonist-evoked Ca2+ release and/or entry for each individual cell as well as
        the maximum peaks, slopes.
        Additionaly, The average values and standard errors for each parameters from the whole experiment will be determined.        
        Moreover, the program will generate a second worksheet, containing the same data than above but normalized to F/F0.
        It will also generates two scattercharts, one on each worksheet, including the traces from all the analyzed cells.
        Finally, the program will create a copy of your file/files with the calculations and the tag '_analyzed' at the end,
        leaving your original file unmodified. 
            
        TL;DR:
            You would need to provide ONE ROW NUMBER (Basal pre-stimuli) or TWO ROW NUMBERS (Basal pre-stimuli and basal pre-calcium adition).
            You might manually modify the adquisiton time (2s), AUC time (150s), slopes duration (30s).
            Results: The AUC and slopes for single cells and average of the whole experiment both in ratio and ratio F/F0.
            
    <<  Disclaimer: Calcium Imaging Analyzer will always read the files, create a copy and keep the original one unmodified.
        Nevertheless, I utterly recommend to work with backed up files, and not original ones. >> 
    
    If you have suggestions, doubts or detect any bug, please contact me at ijp@unex.es.
    Have fun.!
"""


welcome_message_v_1_69 = """
    Welcome to Calcium Imaging Analyzer (CaImAn) v.1.69
    By Isaac Jardin.

    By running this program you will be able to automatically analyze your calcium imaging experiments
    recorded with the Nikon Imaging Software (NIS) or to format them in order to be able to use CaImAn.

    1 - Analysis:
            > Calcium entry experiments (Medium with calcium -> Agonist). Single and multiple files.
            > SOCE experiments (Calcium-free medium [EGTA] -> Agonist -> Extracellular calcium). Single and multiple files.
            > Calcium oscillation experiments (Medium with calcium -> Agonist). Single files.
        
        Formatting:
            > .csv files generated with the Fiji ImageJ macro "Reanálisis_Fura2_Nikon_FINAL_V_2.ijm" modified from Pedro Camello.
            > .xlsx files whose filenames contains spaces (" "), which will be replaced by "_".

        Future upgrades:
            > Calcium oscillation experiments (Medium with calcium -> Agonist). Multiple files.
            > Formating Nikon confocal files.
            > Formating Metafluor files.
    
    2 - You would need to provide ONE ROW NUMBER (Basal pre-stimuli) or TWO ROW NUMBERS (Basal pre-stimuli and basal pre-calcium adition).
        You might manually modify the adquisiton time (2s), AUC time (150s), slopes duration (30s).
        Results: The AUC and slopes for single cells and average of the whole experiment both in ratio and ratio F/F0.
            
    <<  Disclaimer: CaImAn will always read a file, create a copy and keep the original unmodified.
        Nevertheless, I would utterly recommend you to work with backed up files, and not original ones. >> 
    
    If you have suggestions, doubts or detect any bug, please contact me at isaac.jardin@gmail.com.
    Have fun.!
"""

welcome_message_v_1_73 = """
    Welcome to Calcium Imaging Analyzer (CaImAn) v.1.73
    By Isaac Jardin.

    By running this program you will be able to automatically analyze your calcium imaging experiments
    recorded with the Nikon Imaging Software (NIS) or to format them in order to be able to use CaImAn.

    1 - Analysis:
            > Calcium entry experiments (Medium with calcium -> Agonist). Single and multiple files.
            > SOCE experiments (Calcium-free medium [EGTA] -> Agonist -> Extracellular calcium). Single and multiple files.
            > Calcium oscillation experiments (Medium with calcium -> Agonist). Single files and multiple files.
        
        Formatting:
            > .csv files generated with the Fiji ImageJ macro "Reanálisis_Fura2_Nikon_FINAL_V_2.ijm" modified from Pedro Camello.
            > .xlsx files whose filenames contains spaces (" "), which will be replaced by "_".      
    
    2 - You would need to provide ONE ROW NUMBER (Basal pre-stimuli) or TWO ROW NUMBERS (Basal pre-stimuli and basal pre-calcium adition).
        You might manually modify the adquisiton time (2s), AUC time (150s), slopes duration (30s).
        Results: The AUC and slopes for single cells and average of the whole experiment both in ratio and ratio F/F0.
            
    <<  Disclaimer: CaImAn will always read a file, create a copy and keep the original unmodified.
        Nevertheless, I would utterly recommend you to work with backed up files, and not original ones. >> 
    
    If you have suggestions, doubts or detect any bug, please contact me at isaac.jardin@gmail.com.
    Have fun.!
"""


time_initial_linregress_text = """
    To work correctly, this program needs to calculate the fura2 fluorescence initial curve fitting (regression analysis).
    To achieve that, it will use a range of values from 1:X, where:
        
        > 1 is the first point of the trace.
        > X is the last point of the range.
        >> For instance, if you introduce '20', it will perform the regression analysis over the first 20 values of a given trace.

    Please introduce the LAST POINT (X) of your desire range (minimum 15): """

time_final_linregress_text = """
    To work correctly, this program needs to calculate the fura2 fluorescence final curve fitting (regression analysis).
    To achieve that, it will use a range of values from X:Z, where:
        
        > X is the first point of the range.
        > Z is the last point of the trace.
        >> For instance, if you introduce '20', it will perform the regression analysis over the last 20 values of a given trace.

    Please introduce the FIRST POINT (X) of your desire range (minimum 15): """


imageJ_csv_to_xlsx_message = """
    csv_2_xlsx_converter.py -- Converts the csv files generated by the ImageJ - Fiji macro "Reanálisis_Fura2_Nikon_FINAL_V_2" to xlsx files.
    Moreover, it formats the new xlsx file to a elegible file to be analyzed by Calcium_Imaging_Analyzer.py.
    """

filename_without_spaces_message = """
    filename_without_spaces.py -- Generates a copy of .xlsx files whith their filename without spaces,
                  'MDA MB 231 shOrai1' ---> 'MDA_MB_231_shOrai1'
    becoming a elegible file to be analyzed by Calcium_Imaging_Analyzer.py.
    """


integral_time_message = """Set the number of frames(values) taken to calculate the AUC.
By default, 75 frames, which corresponds to 1 frame/2 seconds for 150 seconds.
Leave it empty to set the default value: """


slope_time_release_message = """Set the number of frames(values) taken to calculate the calcium release trendline.
By default, 15 frames, which corresponds to 1 frame/2 seconds for 30 seconds.
Leave it empty to set the default value: """

slope_time_entry_message = """Set the number of frames(values) taken to calculate the calcium entry trendline.
By default, 15 frames, which corresponds to 1 frame/2 seconds for 30 seconds.
Leave it empty to set the default value: """

keyword_message = """Introduce one word with the studied parameter (for example 'Ratio', 'FITC', 'YFP', ...)
or leave it empty to set the default (Ratio): """

analysis_study_dict = ["SOCE analysis (single file).",
                       "Calcium entry analysis (single file).",
                       "Imaging calcium oscillations analysis (single file).",
                       "Confocal calcium oscillations analysis (single file).",
                       "SOCE analysis (multiple files).",
                       "Calcium entry analysis (multiple files).",
                       "Imaging calcium oscillations analysis (multiple files).",
                       "Go back to the main menu.",
                       ]
