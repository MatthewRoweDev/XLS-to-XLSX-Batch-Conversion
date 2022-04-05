# XLS to XLSX Batch Conversion
A PowerShell script I wrote for an answer on r/Excel. I couldn't find any batch XLS>XLSX conversion scripts anywhere else online so I had to do a lot of the research and write it from scratch.

# How to Use (THESE STEPS ARE REQUIRED FOR THE SCRIPT TO WORK)
1. Verify that excelcnv.exe is present in 'C:\Program Files\Microsoft Office\root\Office16\'
   * The default has been set for Office 16, but the location of excelcnv.exe might be different based on which Office version you are using. Usually, the file can be found in the same folder as excel.exe.
   * If excelcnv.exe is NOT present in the default folder for $directoryExcelRoot, change $directoryExcelRoot to point whatever folder does has excelcnv.exe (enclose the path in single-quotation marks).
   * Do NOT move excelcnv.exe, it has to be in its default directory to work.
2. Change $directorySource to the directory that contains the .xls files you want to convert
3. Paste target directories in $directoryOne, $directoryTwo, and $directoryThree.
   * If you plan on expanding to more than three directories, the target directory variables should be added below the declaration of $directoryThree.
   * You will also need to add another switch statement for any additional variables.
