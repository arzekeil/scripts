<#
.SYNOPSIS
    Script to group media files (images, videos, music, documents) by type and modification date, and log skipped files.

.DESCRIPTION
    This PowerShell script organizes files from the input folder (including subfolders) into three main categories:
    "images-and-videos", "music", and "documents", then groups the files within each category by year and month of their modification date.
    Files that don't match these categories are logged into a "logs" folder, and a progress bar is shown during execution.

.AUTHOR
    Your Name (arzekeil)

.DATE_CREATED
    August 18, 2024

.LAST_MODIFIED
    August 18, 2024

.EXAMPLES
    .\Group-Media.ps1
    .\Group-Media.ps1 -InputFolder "C:\path\to\input"
    .\Group-Media.ps1 -InputFolder "C:\path\to\input" -OutputFolder "C:\path\to\output"
    .\Group-Media.ps1 -Help

.PARAMETER InputFolder
    The folder containing media files to be grouped, including files in subfolders. Default: C:\Users\$env:USERNAME\Documents\unsorted

.PARAMETER OutputFolder
    The folder where grouped media files will be placed. Default: C:\Users\$env:USERNAME\Documents\media

.PARAMETER Help
    Displays the help message.

.NOTES
    - Ensure you have the appropriate permissions to access the specified folders.
    - The script moves files, so ensure the destination folder has sufficient space.

#>

param (
    [string]$InputFolder = "C:\Users\$env:USERNAME\Documents\unsorted",
    [string]$OutputFolder = "C:\Users\$env:USERNAME\Documents\media",
    [switch]$Help
)

# Display Help if -Help flag is used
if ($Help) {
    Write-Host "Usage:"
    Write-Host "  Group-Media.ps1 [-InputFolder <path>] [-OutputFolder <path>] [-Help]"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -InputFolder     Specify the input folder path. Default: C:\Users\$env:USERNAME\Documents\unsorted"
    Write-Host "  -OutputFolder    Specify the output folder path. Default: C:\Users\$env:USERNAME\Documents\media"
    Write-Host "  -Help            Show this help message."
    exit
}

# Ensure the output folder exists
if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force
}

# Create logs folder
$logsFolder = Join-Path -Path $OutputFolder -ChildPath "logs"
if (-not (Test-Path $logsFolder)) {
    New-Item -Path $logsFolder -ItemType Directory
}

# Create the log file with the current date and time (24-hour format)
$logFileName = "skipped-files-{0:yyyy-MM-dd_HH-mm-ss}.log" -f (Get-Date)
$logFilePath = Join-Path -Path $logsFolder -ChildPath $logFileName

# Function to group files by type and year/month, and log skipped files
function Group-MediaFiles {
    param (
        [string]$SourceFolder,
        [string]$DestinationFolder
    )

    # File extensions categorized as images and videos
    $imageVideoExtensions = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".mp4", ".mov", ".avi", ".mkv", ".wmv")
    # File extensions categorized as music
    $musicExtensions = @(".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma", ".m4a")
    # File extensions categorized as documents
    $documentExtensions = @(".doc", ".docx", ".pdf", ".txt", ".xls", ".xlsx", ".ppt", ".pptx", ".odt")

    # Get all files to process and initialize counters
    $files = Get-ChildItem -Path $SourceFolder -File -Recurse
    $totalFiles = $files.Count
    $currentFileIndex = 0

    # Process each file and show the progress bar
    $files | ForEach-Object {
        $currentFileIndex++

        # Update the progress bar
        $percentComplete = [math]::Round(($currentFileIndex / $totalFiles) * 100)
        Write-Progress -Activity "Processing Files" -Status "Processing file $currentFileIndex of $totalFiles" -PercentComplete $percentComplete

        $extension = $_.Extension.ToLower()
        $modifiedDate = $_.LastWriteTime
        $yearMonthFolder = "$($modifiedDate.ToString('yyyy-MM'))"

        # Determine the destination based on file type
        if ($imageVideoExtensions -contains $extension) {
            $mediaCategory = "images-and-videos"
        } elseif ($musicExtensions -contains $extension) {
            $mediaCategory = "music"
        } elseif ($documentExtensions -contains $extension) {
            $mediaCategory = "documents"
        } else {
            # Log skipped files
            $logEntry = "{0}, {1}" -f $_.FullName, $extension
            Add-Content -Path $logFilePath -Value $logEntry
            return
        }

        # Create the category folder (images-and-videos, music, or documents) if it doesn't exist
        $categoryFolder = Join-Path -Path $DestinationFolder -ChildPath $mediaCategory
        if (-not (Test-Path $categoryFolder)) {
            New-Item -Path $categoryFolder -ItemType Directory
        }

        # Create the year-month folder inside the category folder
        $targetFolder = Join-Path -Path $categoryFolder -ChildPath $yearMonthFolder
        if (-not (Test-Path $targetFolder)) {
            New-Item -Path $targetFolder -ItemType Directory
        }

        # Move the file to the appropriate folder
        $destinationFile = Join-Path -Path $targetFolder -ChildPath $_.Name
        Move-Item -Path $_.FullName -Destination $destinationFile -Force
    }
}

# Start grouping the media files
Group-MediaFiles -SourceFolder $InputFolder -DestinationFolder $OutputFolder

Write-Host "Media files grouped successfully. Skipped files are logged at: $logFilePath"
