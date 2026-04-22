$ErrorActionPreference = "Stop"

$taskName = "MouseSensitivity"
$scriptPath = "C:\Apps\windows_mouse_settings\mouse_sensitivity.ps1"

if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
    throw "Runtime script not found: $scriptPath`nRun 1_setup_mouse_sensitivity.ps1 first."
}

$arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument $arguments

$triggers = @(
    New-ScheduledTaskTrigger -AtStartup
    New-ScheduledTaskTrigger -AtLogOn
)

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew

try {
    $null = Get-ScheduledTask -TaskName $taskName -ErrorAction Stop
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}
catch {
}

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $triggers `
    -Settings $settings `
    -Description "Run mouse_sensitivity.ps1 at Windows startup and user logon to reapply the mouse sensitivity setting." `
    -RunLevel Highest `
    -Force | Out-Null

Write-Host "Registered scheduled task $taskName"
Write-Host "Script: $scriptPath"
Write-Host "Arguments: $arguments"
Write-Host "Triggers: AtStartup, AtLogOn"
