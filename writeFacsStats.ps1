# Note: if script fails, kill orphan Excel processes via: taskkill /F /IM excel.exe

# Import CSV data
$csvPath = [IO.Path]::Combine((Get-Location).path, "input.csv")
$csv = Import-Csv -Path $csvPath

# Initialize and hide Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Add subset of CSV data as a new sheet in the workbook with special formatting
function Import-Csv-Data {
    param (
        $Csv,
        $Workbook,
        $Name
    )

    # Find the last sheet for correct ordering on insert
    $worksheetCount = $Workbook.Worksheets.Count
    $lastWorksheet = $Workbook.Worksheets.Item($worksheetCount)

    # Insert the new sheet after the last sheet
    $newSheet = $Workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastWorksheet)
    $newSheet.Name = $Name

    # Extract and set column headers preserving order
    $columnHeaders = $Csv[0].PSObject.Properties.Name
    $row = 1
    $column = 1
    foreach ($header in $columnHeaders) {
        $newSheet.Cells.Item($row, $column).Value2 = $header
        $column++
    }

    # Extract and set rows
    $row = 2
    foreach ($entry in $Csv) {
        $column = 1
        foreach ($header in $columnHeaders) {
            $newSheet.Cells.Item($row, $column).Value2 = $entry.$header
            $column++
        }
        $row++
    }

    # Set hard-coded column widths
    foreach ($key in $columnWidths.Keys) {
        $columnIndex = [Array]::IndexOf($columnHeaders, $key) + 1
        $newSheet.Columns.Item($columnIndex).ColumnWidth = [double]$columnWidths[$key]
    }

    # Heighten first row and wrap text so headers are visible
    $rowRange = $newSheet.Rows.Item(1)
    $rowRange.RowHeight = 100
    $rowRange.WrapText = $true

    # Filter sheet
    $startCell = $newSheet.Cells.Item(1, 1)
    $endCell = $newSheet.Cells.Item($row - 1, $columnHeaders.Count)
    $filterRange = $newSheet.Range($startCell, $endCell)
    $filterRange.AutoFilter()
}

function Format-Range-Table {
    param (
        $Range,
        $Worksheet,
        $Workbook,
        $Title,
        $Criteria,
        $Condition = "Condition",
        $filteredCsv,
        $CreateTabs = $False
    )

    # Create thin border around table
    $range = $Worksheet.Range($Range)
    $range.BorderAround(1, 2)

    # Format table title and headers
    $topRange = $range.Rows.Item(1)
    $topRange.Merge($True)
    $topRange.Value2 = $Title
    $topRange.Font.Bold = $True
    $topRange.HorizontalAlignment = -4108

    # Format and set table row names
    $headerRange = $range.Rows.Item(2)
    $headerRange.Font.Bold = $True
    $headerRange.Cells.Item(1, 1) = $Condition
    $headerRange.Cells.Item(1, 2) = "AvgA"
    $headerRange.Cells.Item(1, 3) = "MedA"
    $headerRange.Cells.Item(1, 4) = "Count"
    $headerRange.Cells.Item(1, 5) = "M"
    $headerRange.Cells.Item(1, 6) = "F"

    $global:i = 0 # Use global context due to function scoping in foreach
    foreach ($key in $Criteria.keys) {
        $currentCriteria = $criteria[$key]

        # Set condition name
        $global:i++
        $range.Cells.Item($global:i + 2, 1) = $key

        # Get only CSV entries matching the passed filter criteria
        $filteredEntries = $csv | Where-Object {
            $match = $true
            foreach ($column in $currentCriteria.Keys) {
                if ($_.($column) -ne $currentCriteria[$column]) {
                    $match = $false
                    break
                }
            }
            $match
        }

        # Calculate the average age for the filtered entries
        $averageAge = [Math]::Round(($filteredEntries | Measure-Object -Property "Age" -Average).Average, 2)

        # Order the ages of the filtered entries
        $ages = $filteredEntries | ForEach-Object { [int]$_."Age" } | Sort-Object

        # Calculate the median age from the ordered ages
        $count = $ages.Count
        $medianAge = 0
        if ($count -gt 0) {
            if ($count % 2 -eq 1) {
                $medianAge = $ages[$count / 2 - 0.5]
            } else {
                $medianAge = ($ages[$count / 2 - 1] + $ages[$count / 2]) / 2
            }
        }

        # Get the count of the filtered entries by sex
        $maleCount = ($filteredEntries | Where-Object { $_.Sex -eq "M" }).Count
        $femaleCount = ($filteredEntries | Where-Object { $_.Sex -eq "F" }).Count

        # Fill the table with calculated values
        $range.Cells.Item($global:i + 2, 2) = $averageAge
        $range.Cells.Item($global:i + 2, 3) = $medianAge
        $range.Cells.Item($global:i + 2, 4) = $count
        $range.Cells.Item($global:i + 2, 5) = $maleCount
        $range.Cells.Item($global:i + 2, 6) = $femaleCount

        # Optionally, create a seperate tab for the filtered data
        if ($createTabs) {
            Import-Csv-Data -Csv $filteredEntries -Workbook $Workbook -Name $key
        }
    }
}

# Hardcode widths of certain columns in px
$columnWidths = @{
    "First Name" = 11.5
    "Last Name" = 14
    "Date of study" = 12.5
    "Impression" = 33
    "Endoscopy Impressions" = 33
    "Surgical Pathology Report Diagnosis" = 33
    "Manometry Diagnosis" = 33
    "Bravo pH Monitoring Impressions" = 33
    "Broad Category (choice=Normal)" = 10.5
    "Broad Category (choice=Inconclusive)" = 10.5
    "Broad Category (choice=No Pathology)" = 10.5
    "Broad Category (choice=GERD)" = 10.5
    "Broad Category (choice=NERD)" = 10.5
    "Broad Category (choice=EoE)" = 10.5
    "Broad Category (choice=Reflux Esophagitis)" = 10.5
    "Broad Category (choice=Lymphocytic Esophagitis)" = 10.5
    "Broad Category (choice=Barretts Esophagus)" = 10.5
    "Broad Category (choice=EAC)" = 10.5
    "Intestinal Metaplasia" = 10.5
    "Intestinal Metaplasia Type" = 12.5
    "Mucosa" = 15.5
    "EoE Status" = 9.5
    "Reflux Esophagitis Status" = 11.5
    "Segment Type" = 13.5
    "Dysplasia Status" = 15.5
}

# Setup new workbook
$workbook = $excel.Workbooks.Add()
$newFilePath = [IO.Path]::Combine((Get-Location).path, "output.xlsx")
$workbook.SaveAs($newFilePath)

# Create tables sheet
$worksheet = $workbook.Worksheets.Item(1)
$worksheet.Name = "Tables"

# Create initial data sheet with all CSV data
Import-Csv-Data -Csv $csv -Workbook $workbook -Name "All"

# Create tables and broad category sheets
Format-Range-Table -Range "A1:F10" -Workbook $workbook -Worksheet $worksheet -Title "Samples by Broad Category" -CreateTabs $True -Criteria ([ordered]@{
    "Normal" = @{
        "Broad Category (choice=Normal)" = "Checked"
    }
    "Inconclusive" = @{
        "Broad Category (choice=Inconclusive)" = "Checked"
    }
    "No Pathology" = @{
        "Broad Category (choice=No Pathology)" = "Checked"
    }
    "NERD" = @{
        "Broad Category (choice=NERD)" = "Checked"
    }
    "EoE" = @{
        "Broad Category (choice=EoE)" = "Checked"
    }
    "Reflux Esophagitis" = @{
        "Broad Category (choice=Reflux Esophagitis)" = "Checked"
    }
    "Barrett's Esophagus" = @{
        "Broad Category (choice=Barretts Esophagus)" = "Checked"
    }
    "EAC" = @{
        "Broad Category (choice=EAC)" = "Checked"
    }
})

# Subcategorize Inconclusive Samples
Format-Range-Table -Range "A12:F15" -Workbook $workbook -Worksheet $worksheet -Title "Inconclusive Samples" -Criteria ([ordered]@{
    "Salmon Colored" = @{
        "Broad Category (choice=Inconclusive)" = "Checked"
        "Mucosa" = "Salmon Colored"
    }
    "Intestinal Metaplasia" = @{
        "Broad Category (choice=Inconclusive)" = "Checked"
        "Intestinal Metaplasia" = "IM"
    }
})
Format-Range-Table -Range "H12:M15" -Workbook $workbook -Worksheet $worksheet -Title "Buried Intestinal Metaplasia Samples" -Criteria ([ordered]@{
    "Buried" = @{
        "Broad Category (choice=Inconclusive)" = "Checked"
        "Intestinal Metaplasia" = "IM"
        "Intestinal Metaplasia Type" = "Buried"
    }
    "Non-Buried" = @{
        "Broad Category (choice=Inconclusive)" = "Checked"
        "Intestinal Metaplasia" = "IM"
        "Intestinal Metaplasia Type" = "Non-Buried"
    }
})

# Subcategorize EoE Samples
Format-Range-Table -Range "A17:F20" -Workbook $workbook -Worksheet $worksheet -Title "EoE Samples" -Criteria ([ordered]@{
    "Active" = @{
        "Broad Category (choice=EoE)" = "Checked"
        "EoE Status" = "Active"
    }
    "Remission" = @{
        "Broad Category (choice=EoE)" = "Checked"
        "EoE Status" = "Remission"
    }

})

# Subcategorize Reflux Esophagitis Samples
Format-Range-Table -Range "A22:F27" -Workbook $workbook -Worksheet $worksheet -Title "Reflux Esophagitis Samples" -Criteria ([ordered]@{
    "Active" = @{
        "Broad Category (choice=Reflux Esophagitis)" = "Checked"
        "Reflux Esophagitis Status" = "Active"
    }
    "Treated" = @{
        "Broad Category (choice=Reflux Esophagitis)" = "Checked"
        "Reflux Esophagitis Status" = "Treated"
    }

})
Format-Range-Table -Range "H22:M27" -Workbook $workbook -Worksheet $worksheet -Title "Treated Reflux Esophagitis Samples" -Condition "LA Grade" -Criteria ([ordered]@{
    "A" = @{
        "Broad Category (choice=Reflux Esophagitis)" = "Checked"
        "Reflux Esophagitis Status" = "Treated"
        "LA Grade" = "A"
    }
    "B" = @{
        "Broad Category (choice=Reflux Esophagitis)" = "Checked"
        "Reflux Esophagitis Status" = "Treated"
        "LA Grade" = "B"
    }
    "C" = @{
        "Broad Category (choice=Reflux Esophagitis)" = "Checked"
        "Reflux Esophagitis Status" = "Treated"
        "LA Grade" = "C"
    }
    "D" = @{
        "Broad Category (choice=Reflux Esophagitis)" = "Checked"
        "Reflux Esophagitis Status" = "Treated"
        "LA Grade" = "D"
    }

})
Format-Range-Table -Range "O22:T27" -Workbook $workbook -Worksheet $worksheet -Title "Active Reflux Esophagitis Samples" -Condition "LA Grade" -Criteria ([ordered]@{
    "A" = @{
        "Broad Category (choice=Reflux Esophagitis)" = "Checked"
        "Reflux Esophagitis Status" = "Active"
        "LA Grade" = "A"
    }
    "B" = @{
        "Broad Category (choice=Reflux Esophagitis)" = "Checked"
        "Reflux Esophagitis Status" = "Active"
        "LA Grade" = "B"
    }
    "C" = @{
        "Broad Category (choice=Reflux Esophagitis)" = "Checked"
        "Reflux Esophagitis Status" = "Active"
        "LA Grade" = "C"
    }
    "D" = @{
        "Broad Category (choice=Reflux Esophagitis)" = "Checked"
        "Reflux Esophagitis Status" = "Active"
        "LA Grade" = "D"
    }
})

# Subcategorize Barrett's Esophagus Samples
Format-Range-Table -Range "A29:F34" -Workbook $workbook -Worksheet $worksheet -Title "Barrett's Esophagus Samples" -Criteria ([ordered]@{
    "Active" = @{
        "Broad Category (choice=Barretts Esophagus)" = "Checked"
        "Barrett's Esophagus Status" = "Active"
    }
    "Treated" = @{
        "Broad Category (choice=Barretts Esophagus)" = "Checked"
        "Barrett's Esophagus Status" = "Treated"
    }
})
Format-Range-Table -Range "H29:M34" -Workbook $workbook -Worksheet $worksheet -Title "Active Barrett's Esophagus Samples" -Criteria ([ordered]@{
    "Long Segment" = @{
        "Broad Category (choice=Barretts Esophagus)" = "Checked"
        "Barrett's Esophagus Status" = "Active"
        "Segment Type" = "Long Segment"
    }
    "Short Segment" = @{
        "Broad Category (choice=Barretts Esophagus)" = "Checked"
        "Barrett's Esophagus Status" = "Active"
        "Segment Type" = "Short Segment"
    }
    "Dysplastic" = @{
        "Broad Category (choice=Barretts Esophagus)" = "Checked"
        "Barrett's Esophagus Status" = "Active"
        "Dysplasia Status" = "Dysplastic"
    }
    "Non-Dysplastic" = @{
        "Broad Category (choice=Barretts Esophagus)" = "Checked"
        "Barrett's Esophagus Status" = "Active"
        "Dysplasia Status" = "Non-Dysplastic"
    }
})

# Autofit columns
$worksheet.Columns.Item("A:T").EntireColumn.AutoFit() | Out-Null

# Save and close workbook
$workbook.Save()
$excel.Quit()

# Ensure removal of orphan Excel processes
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
