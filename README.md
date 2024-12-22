# Sort Photos and Videos by Date

A PowerShell script I put together with the help of AI to organize photos and videos into subfolders by year and month, based on the ```Date taken``` or ```Media created``` metadata properties. This script is intended for cleaning up your media library, especially if you have thousands photo and video files from various devices that you wish to organise and catalogue by year and month.

## What it does

1. Prompts for a target directory to process.
2. Reads metadata properties (`Date taken` or `Media created`) from files and parses them.
3. Determines the correct year and month subfolders to be created.
4. Skips files missing `Date taken` or `Media created` properties.
5. Previews the actions (files to move, skipped files) before execution.
6. Prompts to confirm file move.
7. Moves files into `Year\Month` subfolders (e.g., `2023\12` for December 2023).
8. Outputs actions taken into a timestamped log file.

## Requirements

- Windows PowerShell 5.1 or later.
- Access to files with **metadata properties** (`Date taken` or `Media created`).
- Administrative privileges may be required for certain directories.

## Usage

1. Clone the repository: `git clone https://github.com/jackbrumley/media-date-sorter.git`
2. Navigate to the script's directory: `cd media-date-sorter`
3. Run the script: `.\media-date-sorter.ps1`

Follow the on-screen prompts to provide the target directory and confirm actions.

## Screenshots

![Example Screenshot](screenshots/20241222_130747.png)

![Example Screenshot](screenshots/20241222_130845.png)

![Example Screenshot](screenshots/20241222_131630.png)


