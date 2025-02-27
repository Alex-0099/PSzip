# Open Folder Selection Dialog
Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderBrowser.Description = "Select the folder containing subfolders to zip"
$null = $FolderBrowser.ShowDialog()
$SourcePath = $FolderBrowser.SelectedPath

# If no folder is selected, exit
if (!$SourcePath) {
    Write-Host "No folder selected. Exiting..." -ForegroundColor Yellow
    Exit
}

# Ask the user for the preferred format (ZIP or CBZ)
$zipType = Read-Host "Enter format (zip/cbz)"
if ($zipType -ne "zip" -and $zipType -ne "cbz") {
    Write-Host "Invalid choice. Defaulting to zip format." -ForegroundColor Yellow
    $zipType = "zip"
}

# Set Output Path (Same as Source Folder)
$DestinationPath = "$SourcePath\Zipped"

# Ensure destination folder exists
if (!(Test-Path -Path $DestinationPath)) {
    New-Item -ItemType Directory -Path $DestinationPath | Out-Null
}

# Get all folders inside the selected source path
$folders = Get-ChildItem -Path $SourcePath -Directory

foreach ($folder in $folders) {
    $zipFile = "$DestinationPath\$($folder.Name).$zipType"

    # Use Compress-Archive to zip each folder
    Compress-Archive -Path "$($folder.FullName)\*" -DestinationPath $zipFile -Force

    Write-Host "✅ Zipped: $($folder.Name) -> $zipFile"
}

Write-Host "🎉 All folders have been zipped successfully!"