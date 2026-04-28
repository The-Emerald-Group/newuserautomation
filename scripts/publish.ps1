param(
    [string]$Runtime = "win-x64",
    [string]$Configuration = "Release",
    [string]$OutputRoot = "dist",
    [switch]$ZipOutput
)

$ErrorActionPreference = "Stop"

$projectPath = Join-Path $PSScriptRoot "..\NewUserAutomation.App\NewUserAutomation.App.csproj"
$projectPath = (Resolve-Path $projectPath).Path
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

$publishName = "NewUserAutomation-$Runtime"
$publishDir = Join-Path $repoRoot (Join-Path $OutputRoot $publishName)

Write-Host "Publishing $projectPath -> $publishDir" -ForegroundColor Cyan
dotnet publish $projectPath `
  -c $Configuration `
  -r $Runtime `
  --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -o $publishDir

if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed."
}

if ($ZipOutput) {
    $zipPath = Join-Path $repoRoot (Join-Path $OutputRoot "$publishName.zip")
    if (Test-Path $zipPath) {
        Remove-Item $zipPath -Force
    }
    Compress-Archive -Path (Join-Path $publishDir "*") -DestinationPath $zipPath -Force
    Write-Host "Created zip: $zipPath" -ForegroundColor Green
}

Write-Host "Publish complete: $publishDir" -ForegroundColor Green
