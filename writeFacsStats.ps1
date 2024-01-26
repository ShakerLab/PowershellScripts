$csvPath = [IO.Path]::Combine((Get-Location).path, "input.csv")
$csv = Import-Csv -Path $csvPath

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

function Import-Csv-Data {
    param (
        $Csv,
        $Workbook,
        $Name
    )

    $worksheetCount = $Workbook.Worksheets.Count
    $lastWorksheet = $Workbook.Worksheets.Item($worksheetCount)

    $newSheet = $Workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastWorksheet)
    $newSheet.Name = $Name

    $columnHeaders = $Csv[0].PSObject.Properties.Name
    $row = 1
    $column = 1
    foreach ($header in $columnHeaders) {
        $newSheet.Cells.Item($row, $column).Value2 = $header
        $column++
    }

    $row = 2
    foreach ($entry in $Csv) {
        $column = 1
        foreach ($header in $columnHeaders) {
            $newSheet.Cells.Item($row, $column).Value2 = $entry.$header
            $column++
        }
        $row++
    }

    foreach ($key in $columnWidths.Keys) {
        $newSheet.Columns.Item($key).ColumnWidth = [double]$columnWidths[$key]
    }

    $rowRange = $newSheet.Rows.Item(1)
    $rowRange.RowHeight = 100
    $rowRange.WrapText = $true

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

    $range = $Worksheet.Range($Range)
    $range.BorderAround(1, 2)

    $topRange = $range.Rows.Item(1)
    $topRange.Merge($True)
    $topRange.Value2 = $Title
    $topRange.Font.Bold = $True
    $topRange.HorizontalAlignment = -4108

    $headerRange = $range.Rows.Item(2)
    $headerRange.Font.Bold = $True
    $headerRange.Cells.Item(1, 1) = $Condition
    $headerRange.Cells.Item(1, 2) = "AvgA"
    $headerRange.Cells.Item(1, 3) = "MedA"
    $headerRange.Cells.Item(1, 4) = "Count"
    $headerRange.Cells.Item(1, 5) = "M"
    $headerRange.Cells.Item(1, 6) = "F"

    $global:i = 0
    foreach ($key in $Criteria.keys) {
        $currentCriteria = $criteria[$key]

        $global:i++
        $range.Cells.Item($global:i + 2, 1) = $key

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

        $averageAge = [Math]::Round(($filteredEntries | Measure-Object -Property "Age" -Average).Average, 2)

        $ages = $filteredEntries | ForEach-Object { [int]$_."Age" } | Sort-Object

        # Calculate the median age
        $count = $ages.Count
        $medianAge = 0
        if ($count -gt 0) {
            if ($count % 2 -eq 1) {
                $medianAge = $ages[$count / 2 - 0.5]
            } else {
                $medianAge = ($ages[$count / 2 - 1] + $ages[$count / 2]) / 2
            }
        }

        $maleCount = ($filteredEntries | Where-Object { $_.Sex -eq "M" }).Count
        $femaleCount = ($filteredEntries | Where-Object { $_.Sex -eq "F" }).Count

        $range.Cells.Item($global:i + 2, 2) = $averageAge
        $range.Cells.Item($global:i + 2, 3) = $medianAge
        $range.Cells.Item($global:i + 2, 4) = $count
        $range.Cells.Item($global:i + 2, 5) = $maleCount
        $range.Cells.Item($global:i + 2, 6) = $femaleCount

        if ($createTabs) {
            Import-Csv-Data -Csv $filteredEntries -Workbook $Workbook -Name $key
        }
    }
}

$columnWidths = @{
    "C" = 11.5
    "D" = 14
    "E" = 12.5
    "H" = 33
    "I" = 33
    "J" = 33
    "K" = 33
    "L" = 33
    "AM" = 10.5
    "AN" = 10.5
    "AO" = 10.5
    "AP" = 10.5
    "AQ" = 10.5
    "AR" = 10.5
    "AS" = 10.5
    "AT" = 10.5
    "AU" = 10.5
    "AV" = 10.5
    "AW" = 10.5
    "AX" = 12.5
    "AY" = 15.5
    "BA" = 9.5
    "BB" = 11.5
    "BF" = 13.5
    "BG" = 15.5
}

$workbook = $excel.Workbooks.Add()

$newFilePath = [IO.Path]::Combine((Get-Location).path, "output.xlsx")

$workbook.SaveAs($newFilePath)

$worksheet = $workbook.Worksheets.Item(1)
$worksheet.Name = "Tables"

Import-Csv-Data -Csv $csv -Workbook $workbook -Name "All"

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

$workbook.Save()
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
