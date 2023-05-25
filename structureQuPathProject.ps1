# Define project directory
$projectDir = "./"
$projectFile = Join-Path $projectDir "project.qpproj"
$dataDir = Join-Path $projectDir "data"

# Create the data directory if it doesn't exist
if (-not (Test-Path -Path $dataDir)) {
    New-Item -ItemType Directory -Path $dataDir | Out-Null
}

# Get the current highest entryID
$lastID = [int](Get-Content $projectFile | jq '.lastID')

# Find all jpg files
$jpgFiles = Get-ChildItem -Path $projectDir -Filter "*.jpg" -Recurse -File

foreach ($jpgFile in $jpgFiles) {
    # Increment entryID
    ++$lastID

    # Prepare file paths
    $jpgPath = $jpgFile.FullName
    $qpdataPath = $jpgFile.FullName.Replace(".jpg", ".qpdata")

    # Check if qpdata file exists
    if (Test-Path -Path $qpdataPath) {
        # Create new data directory
        $newDataDir = Join-Path $dataDir $lastID
        New-Item -ItemType Directory -Path $newDataDir | Out-Null

        # Copy qpdata file to new data directory
        Copy-Item -Path $qpdataPath -Destination (Join-Path $newDataDir "data.qpdata")
    }

    # Create a new image entry
    $imageEntry = @{
        serverBuilder = @{
            builderType = "uri"
            providerClassName = "qupath.lib.images.servers.bioformats.BioFormatsServerBuilder"
            uri = ("file:/" + $jpgPath) -replace '\\', '/'
            args = @("--series", "0")
        }
        entryID = $lastID
        randomizedName = [guid]::NewGuid().ToString()
        imageName = $jpgFile.Name
        metadata = @{}
    } | ConvertTo-Json -Depth 4

    # Append the new image entry to the project file
    $json = Get-Content $projectFile | ConvertFrom-Json
    $json.images += $imageEntry | ConvertFrom-Json
    $json | ConvertTo-Json -Depth 4 | Set-Content $projectFile

    # Update lastID in the project file
    $jsonContent = Get-Content $projectFile -Raw | jq ".lastID = $lastID"
    $jsonContent | Set-Content $projectFile
}

Write-Output "Processing completed."
