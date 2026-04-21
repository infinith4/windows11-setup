$ErrorActionPreference = "Stop"

$sourceScript = Join-Path $PSScriptRoot "explorer_copytab.ps1"
$targetDirectory = "C:\Apps\windows_explorer_copytab"
$targetScript = Join-Path $targetDirectory "explorer_copytab.ps1"

if (-not (Test-Path -LiteralPath $sourceScript)) {
    throw "Source script not found: $sourceScript"
}

New-Item -ItemType Directory -Force -Path $targetDirectory | Out-Null
Copy-Item -LiteralPath $sourceScript -Destination $targetScript -Force

Write-Host "Copied runtime script to $targetScript"
