function pdf2text {
	param(
		[Parameter(Mandatory=$true)][string]$file
	)
  # Must lead to local copy of itextsharp.dll
	Add-Type -Path "C:\Users\jcastle7\Downloads\temp\ph\itextsharp.dll"
	$pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $file
	$text = ""
	for ($page = 1; $page -le $pdf.NumberOfPages; $page++){
		$text += [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$page)
	}	
	$pdf.Close()
	return $text
}

function ParseDataFromPDF {
	param(
		[Parameter(Mandatory=$true)][string]$filePath
	)

	# Extract the text from the pdf
	$pdfText = pdf2text -file $filePath
	$pdfLines = $pdfText -split "`n"

	# Extract the desired information based on the line
	$genderPhysicianLine = $pdfLines[6] 
	$gender = if ($genderPhysicianLine -match "Gender: (\w+)") { $Matches[1] } else { "" }
	$physician = if ($genderPhysicianLine -match "Physician: (.*),") { $Matches[1] } else { "" }
	
	$referringLine = $pdfLines[8]
	$referredBy = if ($referringLine -match "Referred by: (.*)") { $Matches[1] } else { "" }

	$dobLine = $pdfLines[10]
	$dob = if ($dobLine -match "DOB: (\d{2}/\d{2}/\d{4})") { $Matches[1] } else { "" }

	$patientName = $pdfLines[11].Trim()
	$medicalRecordNumber = $pdfLines[12].Trim()

	$medicationLine = $pdfLines[13]
	$medication = if ($medicationLine -match "Medication: (\w+)") { $Matches[1] } else { "" }

	$dateLine = $pdfLines[14]
	$date = if ($dateLine -match "Date: (\d{2}/\d{2}/\d{4})") { $Matches[1] } else { "" }

	# Create a custom object to hold the data
	$output = New-Object PSObject
	$output | Add-Member -Type NoteProperty -Name "Gender" -Value $gender
	$output | Add-Member -Type NoteProperty -Name "Physician" -Value $physician
	$output | Add-Member -Type NoteProperty -Name "ReferredBy" -Value $referredBy
	$output | Add-Member -Type NoteProperty -Name "DOB" -Value $dob
	$output | Add-Member -Type NoteProperty -Name "PatientName" -Value $patientName
	$output | Add-Member -Type NoteProperty -Name "MedicalRecordNumber" -Value $medicalRecordNumber
	$output | Add-Member -Type NoteProperty -Name "Medication" -Value $medication
	$output | Add-Member -Type NoteProperty -Name "Date" -Value $date
	
	return $output
}

$folderPath = 'C:\path\to\your\pdfs'
$outputPath = 'C:\path\to\your\output.csv'

# Get all pdf files in the specified directory
$pdfFiles = Get-ChildItem -Path $folderPath -Filter *.pdf

# Initialize an array to hold the output data
$outputData = @()

# Process each pdf file
foreach ($file in $pdfFiles) {
	$outputData += ParseDataFromPDF -filePath $file.FullName
}

# Output the data to a csv file
$outputData | Export-Csv -Path $outputPath -NoTypeInformation
