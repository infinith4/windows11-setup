$ErrorActionPreference = "Stop"

$sourceScript = Join-Path $PSScriptRoot "explorer_copytab.ps1"
$targetDirectory = "C:\Apps\windows_explorer_copytab"
$targetScript = Join-Path $targetDirectory "explorer_copytab.ps1"

if (-not (Test-Path -LiteralPath $sourceScript -PathType Leaf)) {
    throw "Source script not found: $sourceScript"
}

New-Item -ItemType Directory -Force -Path $targetDirectory | Out-Null
Copy-Item -LiteralPath $sourceScript -Destination $targetScript -Force

if (-not (Test-Path -LiteralPath $targetScript -PathType Leaf)) {
    throw "Copied file was not created: $targetScript"
}

$sourceInfo = Get-Item -LiteralPath $sourceScript
$targetInfo = Get-Item -LiteralPath $targetScript
$sourceHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $sourceScript).Hash
$targetHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $targetScript).Hash

if ($sourceHash -ne $targetHash) {
    throw "Hash mismatch after copy. Source=$sourceHash Target=$targetHash"
}

Write-Host "Copied latest runtime script to $targetScript"
Write-Host "Source LastWriteTime: $($sourceInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "Target LastWriteTime: $($targetInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Host "SHA256: $targetHash"
