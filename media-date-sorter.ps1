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
  - Exports the final table of all actions to a timestamped CSV file in a "logs" subfolder.
  - Logs all key actions, including errors, skipped files, and user decisions, into a timestamped log file.
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

# Prepare the logs directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$LogDir = Join-Path $ScriptDir "logs"
Ensure-LogsDirectory -Path $LogDir

# Generate timestamp for log and CSV files
$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

# Prepare the log file path
$LogFile = "${Timestamp}_SortLog.log"
$LogPath = Join-Path $LogDir $LogFile

# Prepare the CSV file path
$CsvFile = "${Timestamp}_SortResults.csv"
$CsvPath = Join-Path $LogDir $CsvFile

# Function: Log messages to both console and log file
function Log-Message {
  param (
    [string]$Message,
    [string]$Level = "Info"
  )
  $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $LogEntry = "[$Timestamp][$Level] $Message"

  # Append to the log file
  Add-Content -Path $LogPath -Value $LogEntry

  # Write to the console without timestamp
  Write-Host "[$Level] $Message"
}

# Prompt for the directory if not provided
if (-not $TargetDirectory) {
  Write-Host ""
  $TargetDirectory = Read-Host -Prompt "Enter the target directory"
}

# Validate the directory
if (-not (Test-Path $TargetDirectory)) {
  Log-Message "The specified directory does not exist." "Error"
  exit
}

Log-Message "Script started. Target directory: $TargetDirectory"

# Get all files in the target directory
$Files = Get-ChildItem -Path $TargetDirectory -File
$AllFiles = @()
$SkippedFiles = @()

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
      $SkippedFiles += $File
    }
  } catch {
    $AllFiles += [PSCustomObject]@{
      FileName = $File.FullName
      DateType = "Error"
      Date = "N/A"
      Destination = "N/A"
    }
    Log-Message "Failed to process file: $($File.FullName)" "Error"
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
Log-Message "File list exported to CSV: $CsvPath"

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
      Log-Message "Moved file: $($File.FileName) to $($File.Destination)" "Success"
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
      Log-Message "Moved file: $($File.FileName) to $($File.Destination)" "Success"
    } catch {
      Log-Message "Failed to move file: $($File.FileName)" "Error"
    }
  }
}

# Log skipped files
if ($SkippedFiles.Count -gt 0) {
  foreach ($SkippedFile in $SkippedFiles) {
    Log-Message "Skipped file: $($SkippedFile.FullName). No valid date metadata found." "Warning"
  }
}

Log-Message "Script completed."
Write-Host ""
Write-Host "File sort operation completed." -ForegroundColor Green