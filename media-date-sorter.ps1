<#
.SYNOPSIS
  Sorts files in a target directory based on Windows Shell's "Date taken" 
  or "Media created" property. Files are placed into Year\Month subfolders.

.DESCRIPTION
  - Prompts for a target directory to process.
  - Reads metadata properties (`Date Taken` or `Media Created`) from files.
  - Previews valid and skipped files in a table.
  - Moves valid files into subfolders named Year\Month.
  - Falls back to `Date Modified` for skipped files.
  - Exports the final table of all actions to a timestamped CSV file in a "Logs" subfolder.
#>

param(
    [string]$TargetDirectory
)

# Prompt for the directory if not provided
if (-not $TargetDirectory) {
    # Add a blank line before the prompt
    Write-Host ""
    # Request user input for the target directory
    $TargetDirectory = Read-Host -Prompt "Enter the target directory"
}

# Validate the directory
if (-not (Test-Path $TargetDirectory)) {
    # Exit if the directory does not exist
    Write-Host "The specified directory does not exist." -ForegroundColor Red
    exit
}

# Create a "Logs" subfolder in the script's directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogDir = Join-Path $ScriptDir "Logs"
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

# Prepare the CSV file path
$CsvFile = "SortResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$CsvPath = Join-Path $LogDir $CsvFile

# Get all files in the target directory
$Files = Get-ChildItem -Path $TargetDirectory -File

# Unified array for all files
$AllFiles = @()

# Function to get the index of a file property
function Get-PropertyIndex {
    param (
        [string]$DirectoryPath,
        [string]$PropertyName
    )
    # Access file metadata properties using Shell.Application
    $Shell = New-Object -ComObject Shell.Application
    $Folder = $Shell.Namespace($DirectoryPath)
    for ($Index = 0; $Folder.GetDetailsOf($Folder.Items, $Index) -ne $null; $Index++) {
        if ($Folder.GetDetailsOf($Folder.Items, $Index) -eq $PropertyName) {
            return $Index
        }
    }
    return -1
}

# Function to safely parse and clean date strings
function Try-ParseDate {
    param (
        [string]$DateString
    )
    try {
        # Remove any invisible Unicode characters (e.g., right-to-left marks)
        $CleanedDate = $DateString -replace "[^\x20-\x7E]", ""
        return [datetime]::Parse($CleanedDate)
    } catch {
        # Return null for invalid dates
        return $null
    }
}

# Process each file
foreach ($File in $Files) {
    try {
        # Access file metadata using COM objects
        $Shell = New-Object -ComObject Shell.Application
        $Folder = $Shell.Namespace($File.Directory.FullName)
        $FileObject = $Folder.ParseName($File.Name)

        # Get property indices for "Date taken" and "Media created"
        $DateTakenIndex = Get-PropertyIndex -DirectoryPath $File.Directory.FullName -PropertyName "Date taken"
        $MediaCreatedIndex = Get-PropertyIndex -DirectoryPath $File.Directory.FullName -PropertyName "Media created"

        $DateTaken = $null
        $UsedDateType = ""

        # Attempt to get "Date Taken"
        if ($DateTakenIndex -ge 0) {
            $DateTaken = $Folder.GetDetailsOf($FileObject, $DateTakenIndex)
            $UsedDateType = "Date Taken"
        }

        # If "Date Taken" is not available, attempt "Media Created"
        if (-not $DateTaken -and $MediaCreatedIndex -ge 0) {
            $DateTaken = $Folder.GetDetailsOf($FileObject, $MediaCreatedIndex)
            $UsedDateType = "Media Created"
        }

        # Parse the date or fallback to Date Modified
        $ParsedDate = $null
        if ($DateTaken) {
            $ParsedDate = Try-ParseDate -DateString $DateTaken
        }

        if (-not $ParsedDate -and $UsedDateType -eq "Media Created") {
            $ParsedDate = Try-ParseDate -DateString $DateTaken
        }

        if (-not $ParsedDate) {
            if ($File.LastWriteTime -ne [datetime]"1980-01-01T00:00:00") {
                $ParsedDate = $File.LastWriteTime
                $UsedDateType = "Date Modified"
            } else {
                $ParsedDate = $null
                $UsedDateType = "None Found"
            }
        }

        if ($ParsedDate -ne $null) {
            # Construct year and month directories for sorting
            $Year = $ParsedDate.Year
            $Month = $ParsedDate.ToString('MM')

            $DestinationDir = Join-Path $TargetDirectory "$Year\$Month"
            $DestinationPath = Join-Path $DestinationDir $File.Name

            # Add file details to unified array
            $AllFiles += [PSCustomObject]@{
                FileName = $File.FullName
                DateType = $UsedDateType
                Date = $ParsedDate
                Destination = $DestinationPath
            }
        } else {
            # Add skipped file details to the unified array
            $AllFiles += [PSCustomObject]@{
                FileName = $File.FullName
                DateType = "None Found"
                Date = "N/A"
                Destination = "N/A"
            }
        }
    } catch {
        # Handle and log errors for each file
        $AllFiles += [PSCustomObject]@{
            FileName = $File.FullName
            DateType = "Error"
            Date = "N/A"
            Destination = "N/A"
        }
    }
}

# Display unified table with colour coding
Write-Host ""  # Add blank line before the table
Write-Host "Files Found:" -ForegroundColor White  # Change text to white
Write-Host ""  # Add blank line

# Dynamically calculate column widths
$FileNameWidth = ($AllFiles | ForEach-Object { $_.FileName.Length } | Measure-Object -Maximum).Maximum + 2
$DateTypeWidth = ($AllFiles | ForEach-Object { $_.DateType.Length } | Measure-Object -Maximum).Maximum + 2
$DateWidth = ($AllFiles | ForEach-Object { $_.Date.ToString().Length } | Measure-Object -Maximum).Maximum + 2
$DestinationWidth = ($AllFiles | ForEach-Object { $_.Destination.Length } | Measure-Object -Maximum).Maximum + 2

# Header row
Write-Host ("{0,-$FileNameWidth}{1,-$DateTypeWidth}{2,-$DateWidth}{3,-$DestinationWidth}" -f "File Name", "Date Type", "Date", "Destination")
Write-Host "".PadRight($FileNameWidth + $DateTypeWidth + $DateWidth + $DestinationWidth, '-')

# Output rows with color-coded text
$AllFiles | ForEach-Object {
    $Color = switch ($_.DateType) {
        "Date Taken"     { "Green" }
        "Media Created"  { "Blue" }
        "Date Modified"  { "Cyan" }
        "None Found"     { "Yellow" }
        "Error"          { "Red" }
        default           { "White" }
    }
    Write-Host ("{0,-$FileNameWidth}{1,-$DateTypeWidth}{2,-$DateWidth}{3,-$DestinationWidth}" -f $_.FileName, $_.DateType, $_.Date, $_.Destination) -ForegroundColor $Color
}
Write-Host ""  # Add blank line after the table

# Export table to CSV
$AllFiles | Export-Csv -Path $CsvPath -NoTypeInformation
Write-Host "File list exported to CSV: $CsvPath" -ForegroundColor White  # Change to white

# Confirm and move files based on trusted dates (Date Taken, Media Created)
Write-Host ""  # Add blank line before prompt
Write-Host "Please review the files above or in the exported CSV." -ForegroundColor White
$ConfirmTrusted = Read-Host -Prompt "Would you like to proceed with moving files based on the Date Taken or Media Created dates? (Y/N)"
if ($ConfirmTrusted -eq "Y") {
    foreach ($File in $AllFiles | Where-Object { $_.DateType -eq "Date Taken" -or $_.DateType -eq "Media Created" }) {
        try {
            # Create destination directory if it does not exist
            $DestinationDir = Split-Path -Path $File.Destination -Parent
            if (-not (Test-Path $DestinationDir)) {
                New-Item -ItemType Directory -Path $DestinationDir | Out-Null
            }
            # Move the file
            Move-Item -Path $File.FileName -Destination $File.Destination
        } catch {
            # Handle any errors during the move operation
        }
    }
}

# Confirm and move files based on Date Modified
Write-Host ""  # Add blank line before prompt
$ConfirmModified = Read-Host -Prompt "Would you like to move any file that did not have a Date Taken or Media Created date and move it based on the Date Modified instead? (Y/N)"
if ($ConfirmModified -eq "Y") {
    foreach ($File in $AllFiles | Where-Object { $_.DateType -eq "Date Modified" }) {
        try {
            # Create destination directory if it does not exist
            $DestinationDir = Split-Path -Path $File.Destination -Parent
            if (-not (Test-Path $DestinationDir)) {
                New-Item -ItemType Directory -Path $DestinationDir | Out-Null
            }
            # Move the file
            Move-Item -Path $File.FileName -Destination $File.Destination
        } catch {
            # Handle any errors during the move operation
        }
    }
}

# Log completion of the file sort operation
Write-Host ""  # Add blank line before completion message
Write-Host "File sort operation completed." -ForegroundColor Green
