# Get version from Application Files directory
$appFilesPath = ".\SlingMD.Outlook\publish\Application Files"
$latestVersion = Get-ChildItem $appFilesPath | 
    Where-Object { $_.Name -like "SlingMD.Outlook_*" } | 
    Sort-Object -Property {
        $version = $_.Name -replace "SlingMD\.Outlook_", "" -replace "_", "."
        [version]$version
    } -Descending | 
    Select-Object -First 1

if (-not $latestVersion) {
    Write-Error "No version found in Application Files directory"
    exit 1
}

# Extract version number
$versionNumber = $latestVersion.Name -replace "SlingMD\.Outlook_", "" -replace "_", "."

# Create Releases directory if it doesn't exist
$releasesDir = ".\Releases"
if (-not (Test-Path $releasesDir)) {
    New-Item -ItemType Directory -Path $releasesDir
}

# Create zip file name
$zipFileName = "SlingMD.Outlook_$($versionNumber -replace "\.", "_").zip"
$zipFilePath = Join-Path $releasesDir $zipFileName

# Create a temporary directory for the files we want to include
$tempDir = ".\temp_package"
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}
New-Item -ItemType Directory -Path $tempDir

# Copy only the necessary files
Copy-Item ".\SlingMD.Outlook\publish\SlingMD.Outlook.vsto" $tempDir
Copy-Item ".\SlingMD.Outlook\publish\setup.exe" $tempDir
Copy-Item $latestVersion.FullName "$tempDir\Application Files\$($latestVersion.Name)" -Recurse

# Create the zip file
Compress-Archive -Path "$tempDir\*" -DestinationPath $zipFilePath -Force

# Clean up temp directory
Remove-Item $tempDir -Recurse -Force

Write-Host "Package created successfully: $zipFileName" 