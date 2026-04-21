$ErrorActionPreference = "Stop"

$taskName = "ExplorerCopyTab"
$scriptPath = "C:\Apps\windows_explorer_copytab\explorer_copytab.ps1"
$logPath = Join-Path $env:TEMP "explorer_copytab.log"

if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
    throw "Runtime script not found: $scriptPath`nRun 1_setup_explorer_copytab.ps1 first."
}

$arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -LogPath `"$logPath`""

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument $arguments

$trigger = New-ScheduledTaskTrigger -AtLogOn
$currentUser = if ($env:USERDOMAIN) { "$env:USERDOMAIN\$env:USERNAME" } else { $env:USERNAME }
$principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew

try {
    $null = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}
catch {
}

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description "Launch explorer_copytab.ps1 at logon to duplicate the active Windows Explorer tab with Ctrl+Alt+Shift+D." `
    -Force | Out-Null

Write-Host "Registered scheduled task $taskName"
Write-Host "Script: $scriptPath"
Write-Host "LogPath: $logPath"
Write-Host "Arguments: $arguments"
