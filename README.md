# CaImAn

Calcium Imaging Analyzer (CaImAn) v.1.73 consists of a bunch of scripts written in python and put together to speed the analysis of calcium imaging experiments performed with the NIS-Elements Advanced Research software, from Nikon Corporation.

CaImAn handles single files as well as several files in a folder. The user would only need to provide the path and the required parameter for each experiment. It works with Excel files produced by NIS, or just any other Excel file that mimics the format of those created by NIS.

By introducing some parameters in the console, it will, almost instantly, generate another Excel file containing key data required for the analysis of calcium imaging experiments. In most of the applications, the calculations will be done by Excel, so the generated file would be available for further modifications by the user.

For instance, in classical calcium entry experiments, where cells are maintained in a medium containing calcium chloride (0.3-1.8 mM), CaImAn will calculate the increase of cytosolic calcium concentration upon stimuli as the integral of the rise in cytosolic calcium concentration for 2½ min after the agonist addition. Furthermore, CaImAn will estimate the velocity (slope) of such increase as well as the maximum cytosolic calcium concentration peak.

Currently, CaImAn consist of the following modules:

1 - Analysis:

    Calcium entry experiments (Medium with calcium -> Agonist). Single and multiple files.
    
    SOCE experiments (Calcium-free medium [EGTA] -> Agonist -> Extracellular calcium). Single and multiple files.
    
    Calcium oscillation experiments (Medium with calcium -> Agonist). Single files and multiple files.
    
2 - Formatting:

    .csv files generated with the Fiji ImageJ macro "Reanálisis_Fura2_Nikon_FINAL_V_2.ijm" modified from Pedro Camello.
    
    .xlsx files whose filenames contains spaces (" "), which will be replaced by "_".

DISCLAIMER:

<<< CaImAn does only assist the user with the tedious work of analyzing 20-30 cells in an experiment for +20 experiments/day, by automating the analysis workflow.

This is a little program that I have written for myself and my lab, Ficel, but even when I have extensive experience in the field of imaging calcium analysis, it has only been two months since I started to learn Python (the current date is March 30th, 2021) or any other programming language, beside some short coding in the Fiji ImageJ script editor.

Therefore, use CaImAn under your consideration and double-check your results if you are considering utilizing the obtained data for publication. I do not warrant that the results are accurate, complete, reliable or error free. I do not take any responsibility for any mistakes, misunderstanding or misinterpretation derived from using CaImAn. It is up to the user to understand the sort of analysis that are being performed, to check whether the calculations are correct and finally, to extract a meaning from the obtained data. 

Finally, CaImAn will always read a file and create a copy of it where CaImAn will write the results, keeping the read file unmodified. Nevertheless, I would utterly recommend you to work with backed up files, and not original ones. I do not take any responsibility for any corrupted/lost files that could result from using CaImAn. >>>

If you have suggestions, doubts or detect any bug, please contact me at isaac.jardin@gmail.com.