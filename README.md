# Sort Photos and Videos by Date

A PowerShell script I put together with the help of AI to organize photos and videos into subfolders by year and month, based on the ```Date taken``` or ```Media created``` metadata properties. This script is intended for cleaning up your media library, especially if you have thousands photo and video files from various devices that you wish to organise and catalogue by year and month.

## What it does

1. Prompts for a target directory to process.
2. Reads metadata properties (`Date Taken` or `Media Created`) from files.
3. Attempts to parse the dates, falling back to `Date Modified` if `Date Taken` or `Media Created` are unavailable.
4. Skips files with no valid date metadata (`Date Taken`, `Media Created`, or `Date Modified`).
5. Displays a color-coded table of all files with dates found.
6. Prompts to confirm moving files based on `Date Taken` or `Media Created`.
7. Optionally prompts to move skipped files based on `Date Modified`.
8. Moves files into `Year\Month` subfolders (e.g., `2023\12` for December 2023).
9. Exports a timestamped CSV file documenting all actions (files moved, skipped, etc.) into a `Logs` subfolder.

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


