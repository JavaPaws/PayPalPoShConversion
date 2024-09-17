# Load Excel Interop Assembly
Function Open-Excel {
    Param([string]$csvPath)

    # Create a new Excel application instance
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true

    # Open the CSV file
    $workbook = $excel.Workbooks.Open($csvPath)
    $sheet = $workbook.Sheets.Item(1)

    # Task 1: Delete the contents of cells C1 through E3
    $sheet.Range("C1:E3").ClearContents()

    # Task 2: Move the contents of cell B1 to C1
    $sheet.Cells.Item(1, 3).Value2 = $sheet.Cells.Item(1, 2).Value2
    $sheet.Cells.Item(1, 2).ClearContents()

    # Task 3: Move the contents of cell B3 to C3
    $sheet.Cells.Item(3, 3).Value2 = $sheet.Cells.Item(3, 2).Value2
    $sheet.Cells.Item(3, 2).ClearContents()

    # Task 4: Look for a row where column A and B match SF and USD and copy it to the top
    $lastRow = $sheet.UsedRange.Rows.Count
    for ($row = 1; $row -le $lastRow; $row++) {
        $columnA = $sheet.Cells.Item($row, 1).Value2
        $columnB = $sheet.Cells.Item($row, 2).Value2

        if ($columnA -eq "SF" -and $columnB -eq "USD") {
            $rowToCopy = $row
            $sheet.Rows.Item($row).Copy()
            $sheet.Rows.Item(1).Insert()
            $sheet.Rows.Item(1).PasteSpecial()
            $excel.CutCopyMode = $false
            break
        }
    }

    # Task 4 (continued): Clear all rows below the copied line
    if ($rowToCopy) {
        $rowsToClearStart = $rowToCopy + 1
        #$sheet.Range("A$rowsToClearStart:$lastRow").ClearContents()
        $sheet.Range("A${rowsToClearStart}:A$lastRow").EntireRow.ClearContents()
    }
    
    # Task 5: Delete columns A & B
    $sheet.Columns.Item("A:B").Delete()

    # Task 6: Delete columns G & H
    $sheet.Columns.Item("G:H").Delete()

    # Task 7: Insert a new column to the right of F
    $sheet.Columns.Item("G").Insert()


    # Task 8: Sort data below Row 5 by Column D (now Column C) and then by Column A
    $rangeToSort = $sheet.Range("A6:F$lastRow")  # Define the range below Row 5 to be sorted
    $sort = $sheet.Sort
    $sort.SortFields.Clear()
    $sort.SortFields.Add($sheet.Range("D6:D$lastRow"), 0, 1) # Sort by Column D (now Column C after deletion)
    $sort.SortFields.Add($sheet.Range("A6:A$lastRow"), 0, 1) # Then sort by Column A
    $sort.SetRange($rangeToSort)
    $sort.Header = [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlNo
    $sort.Apply()


    # Task 9: Insert an empty row where the value in Column D changes (starting from Row 6)
    $lastRow = $sheet.UsedRange.Rows.Count
    $prevValue = $sheet.Cells.Item(6, 4).Value2  # Get initial value from Column D in Row 6

    # Iterate from row 7 onward and check where the value in Column D changes
    for ($row = 7; $row -le $lastRow; $row++) {
        $currentValue = $sheet.Cells.Item($row, 4).Value2  # Column D in the current row

        # If the value in Column D changes, insert a new empty row above the current row
        if ($currentValue -ne $prevValue) {
            $sheet.Rows.Item($row).Insert()
            $prevValue = $currentValue
            $row++  # Skip the newly inserted row to avoid infinite loop
            $lastRow++  # Increment lastRow since we added a new row
        }
    }


    # Task 10: Insert description in Column G based on key in Column D
    $mapping = @{}

    # Read the PayPalTransactionCodes.txt file and populate the mapping
    Get-Content $mappingFilePath | ForEach-Object {
        # if ($_ -match '^(.*?):\s(.*)$') {
        if ($_ -match '^(.*?):\s\"(.*)\"$') {
            $mapping[$matches[1].Trim()] = $matches[2].Trim()
        }
    }

    # Iterate through rows starting from Row 6 to the end of the sheet
    for ($row = 6; $row -le $lastRow; $row++) {
        $key = $sheet.Cells.Item($row, 4).Value2  # Column D (the key)
        # Write-Host "Key: Q$key Q"
        
        if ($key -ne $null -and $mapping.ContainsKey($key)) {
            $description = $mapping[$key]
            # Write-Host "Key: $key, Description: $description"
            $sheet.Cells.Item($row, 7).Value2 = "$description"  # Insert the value in Column G
        } elseif ($key -ne $null) {
            $sheet.Cells.Item($row, 7).Value2 = "Description not found"
        }
    }


    # Task 11: Insert Rows at the top for each unique Transaction Event Code
    $usedRange = $sheet.UsedRange    # Get the used range of the sheet

    $uniqueValues = @{}   # Create a hashtable to store unique values from column D

    # Loop through each row starting from row 6
    for ($row = 6; $row -le $usedRange.Rows.Count; $row++) {
        $valueD = $sheet.Cells.Item($row, 4).Value2
        $valueG = $sheet.Cells.Item($row, 7).Value2

        if ($valueD -ne $null -and -not $uniqueValues.ContainsKey($valueD)) {
            $uniqueValues[$valueD] = $valueG
        }
    }

    # Insert rows at the top for each unique value in column D
    $rowIndex = 1
    foreach ($key in $uniqueValues.Keys) {
        $sheet.Rows.Item($rowIndex).Insert([Microsoft.Office.Interop.Excel.XlInsertShiftDirection]::xlShiftDown)
        $sheet.Cells.Item($rowIndex, 6).Value2 = "$key"
        # $sheet.Cells.Item($rowIndex, 7).Value2 = "test"

        if ($key -ne $null -and $mapping.ContainsKey($key)) {
            $description = $mapping[$key]
            $sheet.Cells.Item($rowIndex, 7).Value2 = "$description"
        } 
        $rowIndex++
    }






    # Save changes and open the file in Excel
    # $workbook.Save()
    # $workbook.Close($false)
    
}

# Define the path to the CSV file
$csvFilePath = ".\test.csv"

# Define the path to the mapping file
$mappingFilePath = ".\PayPalTransactionCodes.txt"

# Execute the function to modify and open the file in Excel
Open-Excel -csvPath $csvFilePath -mappingFilePath $mappingFilePath
