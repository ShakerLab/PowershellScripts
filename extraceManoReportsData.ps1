# Note: if script fails, kill orphan Word processes via: taskkill /F /IM winword.exe

$objWord = New-Object -Com Word.Application

enum CaptureTypes {
    None
    Indications
    Findings
    Impression
}

function ParseManoDataFromDoc {
    param (
        [Parameter(Mandatory = $true)][string]$filePath,
        [Parameter(Mandatory = $true)][System.Object]$wordInstance
    )

    $wordDocument = $wordInstance.Documents.Open($filePath)

    if (-not $wordDocument) {
        return
    }
    
    if ($wordDocument.Tables.Count -eq 0) {
        $wordDocument.Close()
        return
    }

    $Table = $wordDocument.Tables.Item(1)

    $patientMrunString = RemoveControlChars $Table.Cell(1, 1).Range.Text

    $pattern = "Patient:(?<LastName>[^,]+),\s*(?<FirstName>[^\d]+)\s+(?<Mrun>\d+)"

    if ($patientMrunString -match $pattern) {
        $mrun = FormatMrun -mrun $matches['Mrun']
        $firstName = FormatName -name $matches['FirstName']
        $lastName = FormatName -name $matches['LastName']
    }
    else {
        $mrun = ''
        $firstName = ''
        $lastName = ''
    }

    $sex = FormatSex -sexString (RemoveControlChars $Table.Cell(1, 3).Range.Text)
    $dob = FormatDate -dateString (RemoveControlChars $Table.Cell(2, 3).Range.Text)
    $reader = FormatName -name (RemoveControlChars $Table.Cell(1, 5).Range.Text)
    $referer = FormatName -name (RemoveControlChars $Table.Cell(3, 5).Range.Text)
    $dos = FormatDate -dateString (RemoveControlChars $Table.Cell(4, 5).Range.Text)

    $indications = ""
    $findings = ""
    $impression = ""

    $capturing = [CaptureTypes]::None

    :exit foreach ($paragraph in $wordDocument.Paragraphs) {
        $paragraphText = RemoveControlChars $paragraph.Range.Text
        #Write-Output $paragraphText

        # Reset capture type on blank line
        if (-not $paragraphText.Trim() -and $capturing -ne [CaptureTypes]::Findings) {
            # Write-Output "RESET"
            $capturing = [CaptureTypes]::None
            continue exit
        }

        switch ($paragraphText) {
            "Indications" {
                $capturing = [CaptureTypes]::Indications
                continue exit
            }
            "Interpretation / Findings" {
                $capturing = [CaptureTypes]::Findings
                continue exit
            }
            "Impressions" {
                $capturing = [CaptureTypes]::Impression
                continue exit
            }
        }


        switch ($capturing) {
            ([CaptureTypes]::Indications) {
                $indications += $paragraphText + "`n"
                continue exit
            }
            ([CaptureTypes]::Findings) {
                $findings += $paragraphText + "`n"
                continue exit
            }
            ([CaptureTypes]::Impression) {
                $impression += $paragraphText + "`n"
                continue exit
            }
        }
    }

    $output = [PSCustomObject]@{
        record_id        = ''
        full_string      = $patientMrunString
        mrun             = $mrun
        first_name       = $firstName
        last_name        = $lastName
        sex              = $sex
        dob              = $dob
        manometry_reader = $reader
        name_referred_md = $referer
        manometry_date   = $dos
        indication       = $indications.TrimEnd("`n")
        findings         = $findings.TrimEnd("`n")
        diagnosis1       = $impression.TrimEnd("`n")
        LastWriteTime    = (Get-Item $filePath).LastWriteTime
    }

    $wordDocument.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordDocument) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    return $output
}

function RemoveControlChars {
    param (
        [Parameter(Mandatory = $true)][string]$controlString
    )

    $output = ($controlString -replace '"', '""' -replace '[\x00-\x1F\x7F]').Trim()
    return $output
}

function FormatName {
    param (
        [Parameter(Mandatory = $true)][string]$name
    )

    # Regular expression pattern to match recognized titles
    $titlePattern = "\b(MD|NP|DDS|DO|PA|RN|DVM|PhD|MS|MA)\b"

    # Capitalize the name part while preserving titles and commas
    $formattedName = ($name -split '(\s+|,)') | ForEach-Object {
        if ($_ -match $titlePattern) {
            $_  # Preserve titles as is
        } elseif ($_ -eq ',') {
            $_  # Preserve commas as is
        } elseif ($_ -ne '') {
            $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower()
        }
    }

    # Join the parts back into a single string
    $formattedName = $formattedName -join ''

    # Return the formatted name
    return $formattedName.Trim()
}

function FormatDate {
    param (
        [Parameter(Mandatory = $true)][string]$dateString
    )

    # Parse the date string using the correct format
    $date = [datetime]::ParseExact($dateString, 'M/d/yyyy', $null)

    # Format the date as YYYY-MM-DD
    $formattedDate = $date.ToString('yyyy-MM-dd')

    return $formattedDate
}

function FormatSex {
    param (
        [Parameter(Mandatory = $true)][string]$sexString
    )

    if ($sexString.StartsWith('M')) {
        return 1
    } else {
        return 2
    }
}

function FormatMrun {
    param (
        [Parameter(Mandatory = $true)][string]$mrun
    )

    # Pad the MRUN with leading zeros to reach a length of 9
    $formattedMrun = $mrun.PadLeft(9, '0')

    return $formattedMrun
}

$folderPath = (Get-Location).path
$outputPath = [IO.Path]::Combine((Get-Location).path, "output.csv")

# Get all pdf files in the specified directory
$wordFiles = Get-ChildItem -Path $folderPath -Filter '*.docx'

# Initialize an array to hold the output data
$entries = @{}

$startingRecordId = [int](Read-Host "Enter the starting record_id value")

# Process each word file
foreach ($file in $wordFiles) {
    Write-Output $file.FullName
    $result = ParseManoDataFromDoc -filePath $file.FullName -wordInstance $objWord

    if ($null -ne $result) {
        $key = "$($result.mrun)-$($result.dos)"

        if (-not $entries.ContainsKey($key) -or ($result.LastWriteTime) -gt ($entries[$key].LastWriteTime)) {
            $entries[$key] = $result
            Write-Output $result
        }
    }
    else {
        Write-Output "No data extracted from $($file.FullName)"
    }
}

# Output the data to a csv file
$outputData = $entries.Values | ForEach-Object {
    $_ | Select-Object -Property record_id, mrun, first_name, last_name, sex, dob, manometry_reader, name_referred_md, manometry_date, indication, findings, diagnosis1
}

$sortedData = $outputData | Sort-Object -Property manometry_date, mrun
$recordId = $startingRecordId
$sortedDataWithId = $sortedData | ForEach-Object {
    $_.record_id = $recordId
    $recordId++
    $_
}

$sortedDataWithId | Export-Csv -Path $outputPath -NoTypeInformation

# Stop Winword Process
$objWord.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objWord) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
