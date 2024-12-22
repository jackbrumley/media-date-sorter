<#
.SYNOPSIS
  Sorts files in a target directory into Year\Month subfolders based on metadata properties.

.DESCRIPTION
  - Prompts for a target directory to process.
  - Reads metadata properties (`Date Taken` or `Media Created`) from files.
  - Falls back to `Date Modified` if no `Date Taken` or `Media Created` properties are found.
  - Skips files with no valid date metadata (`Date Taken`, `Media Created`, or `Date Modified`).
  - Previews all files in a color-coded table, showing dates and destinations.
  - Prompts to confirm moving files with trusted dates (`Date Taken` or `Media Created`).
  - Optionally moves files using `Date Modified` for skipped files.
  - Moves files into `Year\Month` subfolders (e.g., `2023\12` for December 2023).
  - Exports a timestamped CSV file documenting all files and actions into a "logs" subfolder.
#>

param(
  [string]$TargetDirectory
)

# Function: Ensure the logs directory exists
function Ensure-LogsDirectory {
  param ([string]$Path)
  if (-not (Test-Path $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
  }
}

# Function: Get the index of a file property
function Get-PropertyIndex {
  param (
    [string]$DirectoryPath,
    [string]$PropertyName
  )
  $Shell = New-Object -ComObject Shell.Application
  $Folder = $Shell.Namespace($DirectoryPath)
  for ($Index = 0; $Folder.GetDetailsOf($Folder.Items, $Index) -ne $null; $Index++) {
    if ($Folder.GetDetailsOf($Folder.Items, $Index) -eq $PropertyName) {
      return $Index
    }
  }
  return -1
}

# Function: Safely parse and clean date strings
function Try-ParseDate {
  param (
    [string]$DateString
  )
  try {
    $CleanedDate = $DateString -replace "[^\x20-\x7E]", ""
    return [datetime]::Parse($CleanedDate)
  } catch {
    return $null
  }
}

# Function: Log messages
function Log-Message {
  param (
    [string]$Message,
    [string]$Level = "Info"
  )
  $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Write-Host "[$Timestamp][$Level] $Message"
}

# Prompt for the directory if not provided
if (-not $TargetDirectory) {
  Write-Host ""
  $TargetDirectory = Read-Host -Prompt "Enter the target directory"
}

# Validate the directory
if (-not (Test-Path $TargetDirectory)) {
  Write-Host "The specified directory does not exist." -ForegroundColor Red
  exit
}

# Prepare the logs directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogDir = Join-Path $ScriptDir "logs"
Ensure-LogsDirectory -Path $LogDir

# Prepare the CSV file path
$CsvFile = "SortResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$CsvPath = Join-Path $LogDir $CsvFile

# Get all files in the target directory
$Files = Get-ChildItem -Path $TargetDirectory -File
$AllFiles = @()

# Process each file
foreach ($File in $Files) {
  try {
    $Shell = New-Object -ComObject Shell.Application
    $Folder = $Shell.Namespace($File.Directory.FullName)
    $FileObject = $Folder.ParseName($File.Name)

    $DateTakenIndex = Get-PropertyIndex -DirectoryPath $File.Directory.FullName -PropertyName "Date taken"
    $MediaCreatedIndex = Get-PropertyIndex -DirectoryPath $File.Directory.FullName -PropertyName "Media created"

    $DateTaken = $null
    $UsedDateType = ""

    if ($DateTakenIndex -ge 0) {
      $DateTaken = $Folder.GetDetailsOf($FileObject, $DateTakenIndex)
      $UsedDateType = "Date Taken"
    }

    if (-not $DateTaken -and $MediaCreatedIndex -ge 0) {
      $DateTaken = $Folder.GetDetailsOf($FileObject, $MediaCreatedIndex)
      $UsedDateType = "Media Created"
    }

    $ParsedDate = if ($DateTaken) {
      Try-ParseDate -DateString $DateTaken
    } else {
      $null
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
      $Year = $ParsedDate.Year
      $Month = $ParsedDate.ToString('MM')
      $DestinationDir = Join-Path $TargetDirectory "$Year\$Month"
      $DestinationPath = Join-Path $DestinationDir $File.Name

      $AllFiles += [PSCustomObject]@{
        FileName = $File.FullName
        DateType = $UsedDateType
        Date = $ParsedDate
        Destination = $DestinationPath
      }
    } else {
      $AllFiles += [PSCustomObject]@{
        FileName = $File.FullName
        DateType = "None Found"
        Date = "N/A"
        Destination = "N/A"
      }
    }
  } catch {
    $AllFiles += [PSCustomObject]@{
      FileName = $File.FullName
      DateType = "Error"
      Date = "N/A"
      Destination = "N/A"
    }
  }
}

# Display unified table with colour coding
Write-Host ""  
Write-Host "Files Found:" -ForegroundColor White
Write-Host ""

$FileNameWidth = ($AllFiles | ForEach-Object { $_.FileName.Length } | Measure-Object -Maximum).Maximum + 2
$DateTypeWidth = ($AllFiles | ForEach-Object { $_.DateType.Length } | Measure-Object -Maximum).Maximum + 2
$DateWidth = ($AllFiles | ForEach-Object { $_.Date.ToString().Length } | Measure-Object -Maximum).Maximum + 2
$DestinationWidth = ($AllFiles | ForEach-Object { $_.Destination.Length } | Measure-Object -Maximum).Maximum + 2

Write-Host ("{0,-$FileNameWidth}{1,-$DateTypeWidth}{2,-$DateWidth}{3,-$DestinationWidth}" -f "File Name", "Date Type", "Date", "Destination")
Write-Host "".PadRight($FileNameWidth + $DateTypeWidth + $DateWidth + $DestinationWidth, '-')

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
Write-Host ""

# Export table to CSV
$AllFiles | Export-Csv -Path $CsvPath -NoTypeInformation
Write-Host "File list exported to CSV: $CsvPath" -ForegroundColor White

# Confirm and move files based on trusted dates
Write-Host ""
Write-Host "Please review the files above or in the exported CSV." -ForegroundColor White
$ConfirmTrusted = Read-Host -Prompt "Would you like to proceed with moving files based on the Date Taken or Media Created dates? (Y/N)"
if ($ConfirmTrusted -eq "Y") {
  foreach ($File in $AllFiles | Where-Object { $_.DateType -eq "Date Taken" -or $_.DateType -eq "Media Created" }) {
    try {
      $DestinationDir = Split-Path -Path $File.Destination -Parent
      if (-not (Test-Path $DestinationDir)) {
        New-Item -ItemType Directory -Path $DestinationDir | Out-Null
      }
      Move-Item -Path $File.FileName -Destination $File.Destination
    } catch {
      Log-Message "Failed to move file: $($File.FileName)" "Error"
    }
  }
}

# Confirm and move files based on Date Modified
Write-Host ""
$ConfirmModified = Read-Host -Prompt "Would you like to move any file that did not have a Date Taken or Media Created date and move it based on the Date Modified instead? (Y/N)"
if ($ConfirmModified -eq "Y") {
  foreach ($File in $AllFiles | Where-Object { $_.DateType -eq "Date Modified" }) {
    try {
      $DestinationDir = Split-Path -Path $File.Destination -Parent
      if (-not (Test-Path $DestinationDir)) {
        New-Item -ItemType Directory -Path $DestinationDir | Out-Null
      }
      Move-Item -Path $File.FileName -Destination $File.Destination
    } catch {
      Log-Message "Failed to move file: $($File.FileName)" "Error"
    }
  }
}

Write-Host ""
Write-Host "File sort operation completed." -ForegroundColor Green
