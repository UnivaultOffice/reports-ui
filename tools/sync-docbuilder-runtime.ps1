param(
    [string]$BuilderSource = "C:\Source\UnivaultOffice\build_tools\out\win_64\univaultoffice\DocumentBuilder",
    [string]$ReportsUiRoot = $(Resolve-Path (Join-Path $PSScriptRoot "..")),
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $BuilderSource)) {
    throw "Builder source does not exist: $BuilderSource"
}

$reportsUiRootResolved = (Resolve-Path $ReportsUiRoot).Path
$runtimeDir = Join-Path $reportsUiRootResolved "docbuilder\bin"

if ($Clean -and (Test-Path $runtimeDir)) {
    Remove-Item -Path $runtimeDir -Recurse -Force
}

New-Item -ItemType Directory -Path $runtimeDir -Force | Out-Null
Copy-Item -Path (Join-Path $BuilderSource "*") -Destination $runtimeDir -Recurse -Force

Write-Host "DocumentBuilder runtime synced."
Write-Host "Source : $BuilderSource"
Write-Host "Target : $runtimeDir"
