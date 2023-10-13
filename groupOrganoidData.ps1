# Load the CSV file
$csvData = Import-Csv -Path "./measurements.csv"

# Create a hashtable to hold the grouped data
$groupedDataGDR = @{}
$groupedDataGD = @{}
$groupedDataG = @{}

foreach ($row in $csvData) {
    # Calculate Length (um) from Length px
    Add-Member -InputObject $row -NotePropertyName 'Length (um)' -NotePropertyValue ([math]::Round(([double]$row.'Length px' * 1.1941), 2))

    # Skip this row if Length (um) is less than 50
    if ([double]$row.'Length (um)' -lt 50) {
        continue
    }

    # Derive the Group, Day, and Replicate fields from the Image field
    $imageNameParts = $row.Image -split '_'
    Add-Member -InputObject $row -NotePropertyName 'Group' -NotePropertyValue $imageNameParts[0]
    Add-Member -InputObject $row -NotePropertyName 'Day' -NotePropertyValue $imageNameParts[1].substring(1)
    Add-Member -InputObject $row -NotePropertyName 'Replicate' -NotePropertyValue $imageNameParts[2].replace('.jpg', '')

    # Create the column key as a concatenation of Group, Day, and Replicate
    $columnKeyGDR = "$($row.Group)_$($row.Day)_$($row.Replicate)"
    $columnKeyGD = "$($row.Group)_$($row.Day)"
    $columnKeyG = "$($row.Group)"

    # Add the length measurement to the respective group, day, and replicate
    if ($groupedDataGDR.ContainsKey($columnKeyGDR)) {
        $groupedDataGDR[$columnKeyGDR] += $row.'Length (um)'
    }
    else {
        $groupedDataGDR[$columnKeyGDR] = @($row.'Length (um)')
    }

    # Add the length measurement to the respective group and day
    if ($groupedDataGD.ContainsKey($columnKeyGD)) {
        $groupedDataGD[$columnKeyGD] += $row.'Length (um)'
    }
    else {
        $groupedDataGD[$columnKeyGD] = @($row.'Length (um)')
    }

    # Add the length measurement to the respective group
    if ($groupedDataG.ContainsKey($columnKeyG)) {
        $groupedDataG[$columnKeyG] += $row.'Length (um)'
    }
    else {
        $groupedDataG[$columnKeyG] = @($row.'Length (um)')
    }
}

# Determine the maximum number of rows required
$maxRowsGDR = ($groupedDataGDR.Values | Measure-Object -Maximum Count).Maximum
$maxRowsGD = ($groupedDataGD.Values | Measure-Object -Maximum Count).Maximum
$maxRowsG = ($groupedDataG.Values | Measure-Object -Maximum Count).Maximum

# Prepare an array to hold the output rows
$outputRowsGDR = @()
$outputRowsGD = @()
$outputRowsG = @()


# Loop through all three conditions
for ($i = 0; $i -lt $maxRowsGDR; $i++) {
    $outputRow = @{}

    foreach ($columnKeyGDR in $groupedDataGDR.Keys) {
        if ($i -lt $groupedDataGDR[$columnKeyGDR].Count) {
            $outputRow[$columnKeyGDR] = $groupedDataGDR[$columnKeyGDR][$i]
        }
        else {
            $outputRow[$columnKeyGDR] = $null
        }
    }

    $outputRowsGDR += New-Object PSObject -Property $outputRow
}

for ($i = 0; $i -lt $maxRowsGD; $i++) {
    $outputRow = @{}

    foreach ($columnKeyGD in $groupedDataGD.Keys) {
        if ($i -lt $groupedDataGD[$columnKeyGD].Count) {
            $outputRow[$columnKeyGD] = $groupedDataGD[$columnKeyGD][$i]
        }
        else {
            $outputRow[$columnKeyGD] = $null
        }
    }

    $outputRowsGD += New-Object PSObject -Property $outputRow
}

for ($i = 0; $i -lt $maxRowsG; $i++) {
    $outputRow = @{}

    foreach ($columnKeyG in $groupedDataG.Keys) {
        if ($i -lt $groupedDataG[$columnKeyG].Count) {
            $outputRow[$columnKeyG] = $groupedDataG[$columnKeyG][$i]
        }
        else {
            $outputRow[$columnKeyG] = $null
        }
    }

    $outputRowsG += New-Object PSObject -Property $outputRow
}

function Sort-PropertiesAlphabetically {
    param (
        [Parameter(ValueFromPipeline=$true)]
        $InputObject
    )
    
    process {
        $sortedProperties = $InputObject.PSObject.Properties | Sort-Object Name
        $sortedObj = New-Object PSObject
        $sortedProperties | ForEach-Object {
            Add-Member -InputObject $sortedObj -NotePropertyName $_.Name -NotePropertyValue $_.Value
        }
        return $sortedObj
    }
}

# Output the data to a new CSV file
$outputRowsGDR | Sort-PropertiesAlphabetically | Export-Csv -Path "./groupedDataGDR.csv" -NoTypeInformation
$outputRowsGD | Sort-PropertiesAlphabetically | Export-Csv -Path "./groupedDataGD.csv" -NoTypeInformation
$outputRowsG | Sort-PropertiesAlphabetically | Export-Csv -Path "./groupedDataG.csv" -NoTypeInformation
