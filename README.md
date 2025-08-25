Finds all Peaks and Calculates Areas Under the Curve (AUC) for MALDI mass spectrometry imaging analyses.

Utilized in:

Sarikahya et al. 2025 [https://www.nature.com/articles/s41380-025-03113-x]

Cousineau et al. 2025 [https://www.sciencedirect.com/science/article/pii/S0003267025006919]

Ng et al. 2024 [https://www.nature.com/articles/s41386-024-01853-y]

De Felice et al. 2024 [https://www.sciencedirect.com/science/article/pii/S2667174324000740]

Sarikahya et al. 2023 [https://www.nature.com/articles/s41380-023-02190-0]

Sarikahya et al. 2022 [https://www.eneuro.org/content/9/5/ENEURO.0253-22.2022.abstract]

Extensively tested by the Biological Mass Spectrometry and Tissue Imaging Research Lab of Dr. Ken Yeung at Western University
on previously published and non-published datasets to confirm usability across several matrices, including
Norharmane, DPH, FMP-10, ZnO, among others. 

Tested on different computers with MATLAB 2019a and up. MSIreader was used to select regions to generate datasets, but other software also works.

Runtime is approximately 3 minutes to analyze 1000 excel files that contain >25000 m/z values.

Does not effect raw m/z + intensity excel file, will generate a new excel (and .mat) file in a seperate folder that appends _ANALYZED to the file name.
____________________________________________________________________
If there is a specific matrix you use that is not included in the list, feel free to email the author and we will assess its usability with this code.

Please confirm calculated values with this code with personally calculated AUC values ALWAYS before use, as experimental variability in MALDI IMS is a feature of the technique. 
