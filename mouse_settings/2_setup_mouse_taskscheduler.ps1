

$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"C:\Apps\windows_mouse_settings\mouse_sensitivity.ps1`""

$trigger = New-ScheduledTaskTrigger -AtLogOn

Register-ScheduledTask -TaskName "MouseSensitivity" `
    -Action $action `
    -Trigger $trigger `
    -RunLevel Highest `
    -Force