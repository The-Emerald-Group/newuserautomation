param(
    [string]$ZipUrl,
    [string]$AppName = "NewUserAutomation",
    [string]$ExeName = "NewUserAutomation.App.exe",
    [string]$InstallRoot,
    [switch]$NoDesktopShortcut,
    [switch]$BypassBootstrap
)

$ErrorActionPreference = "Stop"
$DefaultZipUrl = "https://repo.emeraldcloud.co.uk/wp-content/Deployment/Emerald%20Applications/tools/NewUserAutomation-win-x64.zip"

function Write-Step {
    param([string]$Message)
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Needs-Elevation {
    param([Parameter(Mandatory = $true)][string]$TargetRoot)

    $fullTarget = [System.IO.Path]::GetFullPath($TargetRoot).TrimEnd('\')
    $programFiles = [System.IO.Path]::GetFullPath($env:ProgramFiles).TrimEnd('\')
    $programFilesX86Raw = ${env:ProgramFiles(x86)}
    $programFilesX86 = if ([string]::IsNullOrWhiteSpace($programFilesX86Raw)) { $null } else { [System.IO.Path]::GetFullPath($programFilesX86Raw).TrimEnd('\') }

    return $fullTarget.StartsWith($programFiles, [System.StringComparison]::OrdinalIgnoreCase) `
        -or ($programFilesX86 -and $fullTarget.StartsWith($programFilesX86, [System.StringComparison]::OrdinalIgnoreCase))
}

function Restart-Elevated {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [string]$ZipUrlArg,
        [string]$AppNameArg,
        [string]$ExeNameArg,
        [string]$InstallRootArg,
        [bool]$NoDesktopShortcutArg
    )

    $argList = @("-ExecutionPolicy", "Bypass", "-File", "`"$ScriptPath`"", "-BypassBootstrap")
    if (-not [string]::IsNullOrWhiteSpace($ZipUrlArg)) { $argList += @("-ZipUrl", "`"$ZipUrlArg`"") }
    if (-not [string]::IsNullOrWhiteSpace($AppNameArg)) { $argList += @("-AppName", "`"$AppNameArg`"") }
    if (-not [string]::IsNullOrWhiteSpace($ExeNameArg)) { $argList += @("-ExeName", "`"$ExeNameArg`"") }
    if (-not [string]::IsNullOrWhiteSpace($InstallRootArg)) { $argList += @("-InstallRoot", "`"$InstallRootArg`"") }
    if ($NoDesktopShortcutArg) { $argList += "-NoDesktopShortcut" }

    Write-Step "Admin rights required. Relaunching installer as administrator..."
    Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList ($argList -join " ")
    exit 0
}

function Restart-BypassExecutionPolicy {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [string]$ZipUrlArg,
        [string]$AppNameArg,
        [string]$ExeNameArg,
        [string]$InstallRootArg,
        [bool]$NoDesktopShortcutArg
    )

    $argList = @("-ExecutionPolicy", "Bypass", "-File", "`"$ScriptPath`"", "-BypassBootstrap")
    if (-not [string]::IsNullOrWhiteSpace($ZipUrlArg)) { $argList += @("-ZipUrl", "`"$ZipUrlArg`"") }
    if (-not [string]::IsNullOrWhiteSpace($AppNameArg)) { $argList += @("-AppName", "`"$AppNameArg`"") }
    if (-not [string]::IsNullOrWhiteSpace($ExeNameArg)) { $argList += @("-ExeName", "`"$ExeNameArg`"") }
    if (-not [string]::IsNullOrWhiteSpace($InstallRootArg)) { $argList += @("-InstallRoot", "`"$InstallRootArg`"") }
    if ($NoDesktopShortcutArg) { $argList += "-NoDesktopShortcut" }

    Write-Step "Relaunching with ExecutionPolicy Bypass..."
    Start-Process -FilePath "powershell.exe" -ArgumentList ($argList -join " ")
    exit 0
}

function Show-InstallProgress {
    param(
        [Parameter(Mandatory = $true)][string]$Activity,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][int]$Percent
    )

    Write-Progress -Activity $Activity -Status $Status -PercentComplete $Percent
}

function Download-FileWithProgress {
    param(
        [Parameter(Mandatory = $true)][string]$SourceUrl,
        [Parameter(Mandatory = $true)][string]$DestinationPath
    )

    $request = [System.Net.HttpWebRequest]::Create($SourceUrl)
    $request.Method = "GET"
    $response = $request.GetResponse()
    $totalBytes = $response.ContentLength
    $responseStream = $response.GetResponseStream()
    $fileStream = [System.IO.File]::Open($DestinationPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)

    try {
        $buffer = New-Object byte[] 81920
        $readTotal = 0L
        while (($read = $responseStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $fileStream.Write($buffer, 0, $read)
            $readTotal += $read

            if ($totalBytes -gt 0) {
                $percent = [int](($readTotal * 100) / $totalBytes)
                $mbDone = "{0:N1}" -f ($readTotal / 1MB)
                $mbTotal = "{0:N1}" -f ($totalBytes / 1MB)
                Show-InstallProgress -Activity "Installing $AppName" -Status "Downloading package ($mbDone MB / $mbTotal MB)" -Percent $percent
            } else {
                Show-InstallProgress -Activity "Installing $AppName" -Status "Downloading package..." -Percent 10
            }
        }
    }
    finally {
        $fileStream.Dispose()
        $responseStream.Dispose()
        $response.Dispose()
    }
}

function New-Shortcut {
    param(
        [Parameter(Mandatory = $true)][string]$ShortcutPath,
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [string]$WorkingDirectory
    )

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    $shortcut.WorkingDirectory = $WorkingDirectory
    $shortcut.IconLocation = "$TargetPath,0"
    $shortcut.Save()
}

if ([string]::IsNullOrWhiteSpace($ZipUrl)) {
    $ZipUrl = $DefaultZipUrl
}
if ([string]::IsNullOrWhiteSpace($ZipUrl) -or $ZipUrl -eq "SET_ME_TO_HOSTED_ZIP_URL") {
    throw "ZipUrl is required. Pass -ZipUrl or set `$DefaultZipUrl in this script."
}

if ([string]::IsNullOrWhiteSpace($InstallRoot)) {
    $InstallRoot = Join-Path $env:LOCALAPPDATA $AppName
}

if (-not $BypassBootstrap) {
    Restart-BypassExecutionPolicy `
        -ScriptPath $PSCommandPath `
        -ZipUrlArg $ZipUrl `
        -AppNameArg $AppName `
        -ExeNameArg $ExeName `
        -InstallRootArg $InstallRoot `
        -NoDesktopShortcutArg ([bool]$NoDesktopShortcut)
}

if ((Needs-Elevation -TargetRoot $InstallRoot) -and -not (Test-IsAdministrator)) {
    Restart-Elevated `
        -ScriptPath $PSCommandPath `
        -ZipUrlArg $ZipUrl `
        -AppNameArg $AppName `
        -ExeNameArg $ExeName `
        -InstallRootArg $InstallRoot `
        -NoDesktopShortcutArg ([bool]$NoDesktopShortcut)
}

$installRoot = $InstallRoot
$currentDir = Join-Path $installRoot "current"
$tempRoot = Join-Path $env:TEMP ("{0}-install-{1}" -f $AppName, [Guid]::NewGuid().ToString("N"))
$zipPath = Join-Path $tempRoot "$AppName.zip"
$extractPath = Join-Path $tempRoot "extracted"

try {
    Write-Step "Preparing temporary workspace"
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $extractPath -Force | Out-Null

    Write-Step "Downloading package"
    Show-InstallProgress -Activity "Installing $AppName" -Status "Starting download..." -Percent 5
    Download-FileWithProgress -SourceUrl $ZipUrl -DestinationPath $zipPath

    Write-Step "Extracting package"
    Show-InstallProgress -Activity "Installing $AppName" -Status "Extracting files..." -Percent 65
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

    $sourceRoot = $extractPath
    $exeCandidates = Get-ChildItem -Path $extractPath -Recurse -Filter $ExeName -File
    if ($exeCandidates.Count -eq 0) {
        throw "Could not find '$ExeName' in downloaded package."
    }

    if ($exeCandidates.Count -ge 1) {
        $sourceRoot = Split-Path -Path $exeCandidates[0].FullName -Parent
    }

    Write-Step "Stopping running app (if open)"
    Get-Process -Name ([System.IO.Path]::GetFileNameWithoutExtension($ExeName)) -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

    Write-Step "Removing previous install at $currentDir"
    Show-InstallProgress -Activity "Installing $AppName" -Status "Removing previous version..." -Percent 78
    if (Test-Path $currentDir) {
        Remove-Item -Path $currentDir -Recurse -Force
    }

    Write-Step "Installing to $currentDir"
    Show-InstallProgress -Activity "Installing $AppName" -Status "Copying files..." -Percent 85
    New-Item -ItemType Directory -Path $currentDir -Force | Out-Null
    Copy-Item -Path (Join-Path $sourceRoot "*") -Destination $currentDir -Recurse -Force

    $exePath = Join-Path $currentDir $ExeName
    if (-not (Test-Path $exePath)) {
        throw "Install completed, but executable was not found at '$exePath'."
    }

    Write-Step "Creating Start Menu shortcut"
    Show-InstallProgress -Activity "Installing $AppName" -Status "Creating shortcuts..." -Percent 92
    $startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
    $startMenuShortcut = Join-Path $startMenuDir "$AppName.lnk"
    New-Shortcut -ShortcutPath $startMenuShortcut -TargetPath $exePath -WorkingDirectory $currentDir

    if (-not $NoDesktopShortcut) {
        Write-Step "Creating Desktop shortcut"
        $desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "$AppName.lnk"
        New-Shortcut -ShortcutPath $desktopShortcut -TargetPath $exePath -WorkingDirectory $currentDir
    }

    Write-Step "Installation complete"
    Show-InstallProgress -Activity "Installing $AppName" -Status "Completed." -Percent 100
    Write-Progress -Activity "Installing $AppName" -Completed
    Write-Host ""
    Write-Host "Installed: $exePath" -ForegroundColor Green
    Write-Host "Start Menu shortcut: $AppName" -ForegroundColor Green
    if (-not $NoDesktopShortcut) {
        Write-Host "Desktop shortcut created." -ForegroundColor Green
    }
    Write-Host ""
    Write-Host "You can now launch $AppName from Start Menu or Desktop."
}
catch {
    Write-Progress -Activity "Installing $AppName" -Completed
    Write-Host ""
    Write-Host "Installation failed: $($_.Exception.Message)" -ForegroundColor Red
    throw
}
finally {
    if (Test-Path $tempRoot) {
        Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($Host.Name -eq "ConsoleHost") {
        Write-Host ""
        Read-Host "Press Enter to close"
    }
}
