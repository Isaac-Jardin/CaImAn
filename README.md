# CaImAn

Calcium Imaging Analyzer (CaImAn) consists of a bunch of scripts written in python and put together to speed the analysis of calcium imaging experiments performed with the NIS-Elements Advanced Research software from Nikon.

CaImAn handles single files as well as several files in a folder. The user would only need to provide the path and the required parameter for each experiment. It works with Excel files produced by NIS, or just any other Excel file with the same format than the ones created by NIS.

Just by introducing some parameters in the console, it will, almost instantly, generate another Excel file containing key data required for the analysis of calcium imaging experiments. In most of the applications, the calculations will be done by Excel, so the generated file is ready to be modified by the user.

For instance, in classical calcium entry experiments, where cells are maintained in a medium containing calcium chloride (0.3-1.8 mM), CaImAn will calculate the increase of cytosolic calcium concentration upon stimuli as the integral of the rise in cytosolic calcium concentration for 2½ min after the addition of the agonist. In addition, CaImAn will estimate the velocity (slope) of such increase as well as the maximum peak of calcium concentration.

Currently, CaImAn consist of the following modules:

1 - Analysis:

    Calcium entry experiments (Medium with calcium -> Agonist). Single and multiple files.
    
    SOCE experiments (Calcium-free medium [EGTA] -> Agonist -> Extracellular calcium). Single and multiple files.
    
    Calcium oscillation experiments (Medium with calcium -> Agonist). Single files.
    
2 - Formatting:

    .csv files generated with the Fiji ImageJ macro "Reanálisis_Fura2_Nikon_FINAL_V_2.ijm" modified from Pedro Camello.
    
    .xlsx files whose filenames contains spaces (" "), which will be replaced by "_".

DISCLAIMER:

<<< CaImAn does only assist the user with the tedious work of analyzing 20-30 cells in an experiment for +20 experiments/day, by automating the analysis workflow. It is up to the user to understand the sort of analysis that are being performed, to check whether the calculations are correct or wrong and finally, to extract a meaning from the obtained data.

This is a little program that I have written for myself and my lab, and even when I have extensive experience in the field of imaging calcium, it has only been two months since I started to learn Python (the current date is March 30th, 2021) or any other programming language, beside some short code in the Fiji ImageJ script editor. Therefore, I invite you to use CaImAn, but please, double-check your results if you are considering using the obtained data for publication.” >>>
