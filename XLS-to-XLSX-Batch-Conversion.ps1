# Declare variables
# Sets variable for the Office root directory. This is required because excelcnv.exe does not work when run from another directory.
$directoryExcelRoot = 'C:\Program Files\Microsoft Office\root\Office16\'

# Sets variable for the source directory that files will be pulled from.
$directorySource = 'D:\Misc\TestDirectory\'

# Sets target directory for each type of file that will be converted and copied copied.
$directoryOne = ''
$directoryTwo = ''
$directoryThree = ''

# Creates array of all .xls files in source directory.
$filesXLS = Get-ChildItem $directorySource *.xls

# Perform the conversion
# Sets console location to Office root directory so excelcnv.exe will work.
Set-Location $directoryExcelRoot

# For each .xls file from the source directory.
foreach ($item in $filesXLS) {
    # Print the name of file that is being converted.
    $item.BaseName
    # Check if file base name contains one of the matching terms
    switch -Wildcard ($item.BaseName) {
        # Each switch statement uses essentially the same algorithm, so only the first switch statement will be documented. The only difference between the algorithms inside each switch is that the $destinationFile variable is built using a different target directory variable based on which term the file name contains.
        # Each switch statement is formatted as "*___*" with the ___ being the term from the file name that you want to search for.
        "*one*" {
            # Creates destination file's full name, stores it in $destinationFile. This full name is made of up the destination directory, the current $item's BaseName, and the string '.xlsx')
            $destinationFile = $directoryOne + $item.BaseName + '.xlsx'
            # Convert file to .xlsx via excelcnv.exe, send to destination file in destination directory
            .\excelcnv.exe -oice $item.FullName $destinationFile
        }
        "*two*" {
            $destinationFile = $directoryTwo + $item.BaseName + '.xlsx'
            .\excelcnv.exe -oice $item.FullName $destinationFile
        }
        "*three*" { 
            $destinationFile = $directoryThree + $item.BaseName + '.xlsx'
            .\excelcnv.exe -oice $item.FullName $destinationFile
        }
        # If the file's base name does not match any of the switch statements, nothing will be done to the file and "No matching terms found" will be printed in the console
        Default {
            Write-Output "No matching terms found"
        }
    }
}

# Move back to source directory
Set-Location $directorySource