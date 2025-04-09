param(
    [Parameter(Mandatory=$false)]
    [string]$Configuration = "Release",
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipPublish = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipClean = $false
)

# Set working directory to script location
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

# Function to write colored output with timestamp
function Write-Log {
    param(
        [string]$Message,
        [System.ConsoleColor]$Color = [System.ConsoleColor]::White
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    Write-Host "[$timestamp] " -NoNewline -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor $Color
}

# Function to find build tool
function Find-BuildTool {
    Write-Log "Searching for build tools..." -Color Gray
    
    # Try dotnet CLI first
    if (Get-Command "dotnet" -ErrorAction SilentlyContinue) {
        Write-Log "Found dotnet CLI" -Color Green
        return "dotnet"
    }
    
    # Try MSBuild through Visual Studio
    try {
        $vsPath = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if ($vsPath) {
            $msbuildPath = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path $msbuildPath) {
                Write-Log "Found MSBuild at: $msbuildPath" -Color Green
                return $msbuildPath
            }
        }
    }
    catch {
        # Continue to next option
    }
    
    Write-Log "No build tools found. Please install .NET SDK or Visual Studio." -Color Red
    return $null
}

# Function to clean the solution
function Start-Clean {
    param(
        [string]$BuildTool
    )
    
    if ($SkipClean) {
        Write-Log "Skipping solution cleaning as requested" -Color Yellow
        return $true
    }
    
    Write-Log "Cleaning solution..." -Color Gray
    
    if ($BuildTool -eq "dotnet") {
        dotnet clean --configuration $Configuration --verbosity minimal
    }
    else {
        & $BuildTool SlingMD.sln /t:Clean /p:Configuration=$Configuration /v:minimal
    }
    
    if ($LASTEXITCODE -ne 0) {
        Write-Log "Clean operation failed with exit code $LASTEXITCODE" -Color Red
        return $false
    }
    
    Write-Log "Clean operation completed successfully" -Color Green
    return $true
}

# Function to build the solution
function Start-Build {
    param(
        [string]$BuildTool
    )
    
    Write-Log "Building solution..." -Color Gray
    
    if ($BuildTool -eq "dotnet") {
        dotnet build --configuration $Configuration --verbosity minimal
    }
    else {
        & $BuildTool SlingMD.sln /t:Build /p:Configuration=$Configuration /v:minimal
    }
    
    if ($LASTEXITCODE -ne 0) {
        Write-Log "Build failed with exit code $LASTEXITCODE" -Color Red
        return $false
    }
    
    Write-Log "Build completed successfully" -Color Green
    return $true
}

# Function to publish the project
function Start-Publish {
    param(
        [string]$BuildTool
    )
    
    Write-Log "Publishing SlingMD..." -Color Gray
    
    if ($BuildTool -eq "dotnet") {
        dotnet publish SlingMD.Outlook\SlingMD.Outlook.csproj --configuration $Configuration --verbosity minimal
    }
    else {
        & $BuildTool SlingMD.Outlook\SlingMD.Outlook.csproj /t:Publish /p:Configuration=$Configuration /v:minimal
    }
    
    if ($LASTEXITCODE -ne 0) {
        Write-Log "Publish operation failed with exit code $LASTEXITCODE" -Color Red
        return $false
    }
    
    # Run package script if it exists
    if (Test-Path ".\package-release.ps1") {
        Write-Log "Running package script..." -Color Gray
        & ".\package-release.ps1"
        
        if ($LASTEXITCODE -ne 0) {
            Write-Log "Package script failed with exit code $LASTEXITCODE" -Color Red
            return $false
        }
    }
    
    Write-Log "Publish completed successfully" -Color Green
    return $true
}

# Main execution
Write-Log "SlingMD Build and Publish Tool" -Color Magenta
Write-Log "----------------------------" -Color Magenta
Write-Log "Configuration: $Configuration" -Color Gray
Write-Log "Skip Clean: $SkipClean" -Color Gray
Write-Log "Skip Publish: $SkipPublish" -Color Gray
Write-Log "----------------------------" -Color Magenta

# Find build tool
$buildTool = Find-BuildTool
if (-not $buildTool) {
    exit 1
}

# Clean solution
$cleanSuccess = Start-Clean -BuildTool $buildTool
if (-not $cleanSuccess -and -not $SkipClean) {
    Write-Log "Clean operation failed. Build process stopped." -Color Red
    exit 1
}

# Build solution
$buildSuccess = Start-Build -BuildTool $buildTool
if (-not $buildSuccess) {
    Write-Log "Build failed. Publish operation skipped." -Color Red
    exit 1
}

# Publish if needed
if (-not $SkipPublish) {
    $publishSuccess = Start-Publish -BuildTool $buildTool
    if (-not $publishSuccess) {
        Write-Log "Publish operation failed." -Color Red
        exit 1
    }
    
    Write-Log "SlingMD was successfully built and published!" -Color Green
} 
else {
    Write-Log "SlingMD was successfully built. Publish was skipped." -Color Green
}

exit 0 