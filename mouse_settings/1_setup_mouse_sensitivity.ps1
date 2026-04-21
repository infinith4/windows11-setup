$script = @'
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class MouseHelper {
    [DllImport("user32.dll")]
    public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);
}
"@
[MouseHelper]::SystemParametersInfo(0x0071, 0, [IntPtr]9, 0x01)
'@

$scriptPath = "C:\Apps\windows_mouse_settings\mouse_sensitivity.ps1"
New-Item -ItemType Directory -Force -Path "C:\Apps\windows_mouse_settings" | Out-Null
$script | Out-File -FilePath $scriptPath -Encoding UTF8
