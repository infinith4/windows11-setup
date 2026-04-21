param(
    [string]$HotKey = "Ctrl+Alt+Shift+D",
    [int]$NewTabDelayMs = 450,
    [int]$AddressBarDelayMs = 150,
    [int]$PostNewTabSettleDelayMs = 250,
    [int]$PostEnterVerifyDelayMs = 500,
    [int]$RunOnceWaitMs = 5000,
    [string]$LogPath = "",
    [switch]$RunOnce
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $env:TEMP "explorer_copytab.log"
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    Add-Content -LiteralPath $LogPath -Value "$timestamp $Message"
}

if (-not ("ExplorerCopyTab.NativeMethods" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ExplorerCopyTab
{
    public static class NativeMethods
    {
        public const int WM_HOTKEY = 0x0312;
        public const uint MOD_ALT = 0x0001;
        public const uint MOD_CONTROL = 0x0002;
        public const uint MOD_SHIFT = 0x0004;
        public const uint MOD_WIN = 0x0008;

        public const uint KEYEVENTF_KEYUP = 0x0002;
        public const uint KEYEVENTF_UNICODE = 0x0004;

        public const ushort VK_CONTROL = 0x11;
        public const ushort VK_SHIFT = 0x10;
        public const ushort VK_MENU = 0x12;
        public const ushort VK_LWIN = 0x5B;
        public const ushort VK_RETURN = 0x0D;
        public const ushort VK_L = 0x4C;
        public const ushort VK_T = 0x54;
        public const ushort VK_A = 0x41;
        public const ushort VK_C = 0x43;
        public const ushort VK_D = 0x44;
        public const ushort VK_ESCAPE = 0x1B;

        [StructLayout(LayoutKind.Sequential)]
        public struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public UIntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public POINT pt;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public uint type;
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;
            [FieldOffset(0)]
            public KEYBDINPUT ki;
            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern sbyte GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport("user32.dll")]
        public static extern bool PeekMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        public static string GetWindowClassName(IntPtr hWnd)
        {
            var builder = new StringBuilder(256);
            int length = GetClassName(hWnd, builder, builder.Capacity);
            return length > 0 ? builder.ToString() : string.Empty;
        }

        public static int WaitForNextHotKeyId()
        {
            MSG message;
            while (true)
            {
                if (PeekMessage(out message, IntPtr.Zero, 0, 0, 1))
                {
                    if (message.message == 0x0012) // WM_QUIT
                        return 0;
                    if (message.message == WM_HOTKEY)
                        return unchecked((int)message.wParam.ToUInt32());
                }
                else
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }

        private static INPUT CreateVirtualKeyInput(ushort vk, bool keyUp)
        {
            return new INPUT
            {
                type = 1,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = vk,
                        wScan = 0,
                        dwFlags = keyUp ? KEYEVENTF_KEYUP : 0,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };
        }

        private static INPUT CreateUnicodeInput(char character, bool keyUp)
        {
            return new INPUT
            {
                type = 1,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = character,
                        dwFlags = KEYEVENTF_UNICODE | (keyUp ? KEYEVENTF_KEYUP : 0),
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };
        }

        public static uint SendKeyPress(ushort vk)
        {
            var inputs = new[]
            {
                CreateVirtualKeyInput(vk, false),
                CreateVirtualKeyInput(vk, true)
            };
            return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        public static uint SendModifiedKeyPress(ushort modifier, ushort vk)
        {
            var inputs = new[]
            {
                CreateVirtualKeyInput(modifier, false),
                CreateVirtualKeyInput(vk, false),
                CreateVirtualKeyInput(vk, true),
                CreateVirtualKeyInput(modifier, true)
            };
            return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        public static void ReleaseModifierKeys()
        {
            var inputs = new[]
            {
                CreateVirtualKeyInput(VK_CONTROL, true),
                CreateVirtualKeyInput(VK_MENU, true),
                CreateVirtualKeyInput(VK_SHIFT, true),
                CreateVirtualKeyInput(VK_LWIN, true),
            };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        public static uint SendUnicodeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            var inputs = new INPUT[text.Length * 2];
            int index = 0;
            foreach (char character in text)
            {
                inputs[index++] = CreateUnicodeInput(character, false);
                inputs[index++] = CreateUnicodeInput(character, true);
            }

            return SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }
    }
}
"@
}

function Resolve-HotKeyDefinition {
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    $normalized = $Value.Trim().ToUpperInvariant()
    $parts = $normalized -split "\+"
    if ($parts.Count -lt 2) {
        throw "HotKey must include at least one modifier and one key. Example: Ctrl+Shift+M"
    }

    $modifiers = [uint32]0
    $keyName = $parts[-1]

    foreach ($part in $parts[0..($parts.Count - 2)]) {
        switch ($part) {
            "CTRL" { $modifiers = $modifiers -bor [ExplorerCopyTab.NativeMethods]::MOD_CONTROL }
            "SHIFT" { $modifiers = $modifiers -bor [ExplorerCopyTab.NativeMethods]::MOD_SHIFT }
            "ALT" { $modifiers = $modifiers -bor [ExplorerCopyTab.NativeMethods]::MOD_ALT }
            "WIN" { $modifiers = $modifiers -bor [ExplorerCopyTab.NativeMethods]::MOD_WIN }
            default { throw "Unsupported modifier in HotKey: $part" }
        }
    }

    if ($keyName.Length -ne 1 -or -not [char]::IsLetterOrDigit($keyName[0])) {
        throw "Only single alphanumeric hotkeys are supported. Current key: $keyName"
    }

    [pscustomobject]@{
        Modifiers = $modifiers
        VirtualKey = [uint32][byte][char]$keyName
    }
}

function Get-ActiveExplorerContext {
    $hwnd = [ExplorerCopyTab.NativeMethods]::GetForegroundWindow()
    if ($hwnd -eq [IntPtr]::Zero) {
        Write-Log "GetForegroundWindow returned zero."
        return $null
    }

    $className = [ExplorerCopyTab.NativeMethods]::GetWindowClassName($hwnd)
    Write-Log "Foreground hwnd=$($hwnd.ToInt64()) class=$className"
    if ($className -notin @("CabinetWClass", "ExploreWClass")) {
        return $null
    }

    $shell = New-Object -ComObject Shell.Application
    foreach ($window in @($shell.Windows())) {
        try {
            $windowHandle = [IntPtr]::new([int64]$window.HWND)
            if ($windowHandle -ne $hwnd) {
                continue
            }

            $path = $window.Document.Folder.Self.Path
            if ([string]::IsNullOrWhiteSpace($path)) {
                Write-Log "Matched Explorer hwnd=$($hwnd.ToInt64()) but path was empty."
                return $null
            }

            Write-Log "Matched Explorer hwnd=$($hwnd.ToInt64()) path=$path title=$([string]$window.LocationName)"
            return [pscustomobject]@{
                Hwnd = $hwnd
                Path = $path
                Title = [string]$window.LocationName
            }
        }
        catch {
            continue
        }
    }

    return $null
}

function Get-ExplorerWindowsForHwnd {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Hwnd
    )

    $shell = New-Object -ComObject Shell.Application
    $matches = @()
    foreach ($window in @($shell.Windows())) {
        try {
            $windowHandle = [IntPtr]::new([int64]$window.HWND)
            if ($windowHandle -ne $Hwnd) {
                continue
            }

            $path = $null
            try {
                $path = $window.Document.Folder.Self.Path
            }
            catch {
                $path = $null
            }

            $matches += [pscustomobject]@{
                Hwnd = $windowHandle
                Path = [string]$path
                Title = [string]$window.LocationName
                Window = $window
            }
        }
        catch {
            continue
        }
    }

    return $matches
}

function Wait-ForActiveExplorerContext {
    param(
        [Parameter(Mandatory)]
        [int]$TimeoutMs
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($stopwatch.ElapsedMilliseconds -lt $TimeoutMs) {
        $context = Get-ActiveExplorerContext
        if ($null -ne $context) {
            return $context
        }

        Start-Sleep -Milliseconds 100
    }

    return $null
}

function Wait-ForWindowForeground {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$Hwnd,
        [int]$TimeoutMs = 800
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($stopwatch.ElapsedMilliseconds -lt $TimeoutMs) {
        if ([ExplorerCopyTab.NativeMethods]::GetForegroundWindow() -eq $Hwnd) {
            return $true
        }

        Start-Sleep -Milliseconds 40
    }

    return $false
}

function Get-ActiveExplorerNavigationTarget {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Context
    )

    $clipboardBackup = $null
    $clipboardHadText = $false
    $sentinel = "__explorer_copytab__{0}" -f ([guid]::NewGuid().ToString("N"))

    try {
        try {
            $clipboardBackup = Get-Clipboard -Raw -Format Text
            $clipboardHadText = $true
        }
        catch {
            $clipboardBackup = $null
            $clipboardHadText = $false
        }

        Set-Clipboard -Value $sentinel
        Start-Sleep -Milliseconds 50

        $ctrlLResult = [ExplorerCopyTab.NativeMethods]::SendModifiedKeyPress(
            [ExplorerCopyTab.NativeMethods]::VK_CONTROL,
            [ExplorerCopyTab.NativeMethods]::VK_L
        )
        if ($ctrlLResult -eq 0) {
            $ctrlLError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Log "Capture source Ctrl+L result=0 Win32Error=$ctrlLError"
        }
        else {
            Write-Log "Capture source Ctrl+L result=$ctrlLResult (expected 4)"
        }

        Start-Sleep -Milliseconds $AddressBarDelayMs

        $ctrlAResult = [ExplorerCopyTab.NativeMethods]::SendModifiedKeyPress(
            [ExplorerCopyTab.NativeMethods]::VK_CONTROL,
            [ExplorerCopyTab.NativeMethods]::VK_A
        )
        if ($ctrlAResult -eq 0) {
            $ctrlAError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Log "Capture source Ctrl+A result=0 Win32Error=$ctrlAError"
        }
        else {
            Write-Log "Capture source Ctrl+A result=$ctrlAResult (expected 4)"
        }

        Start-Sleep -Milliseconds 80

        $ctrlCResult = [ExplorerCopyTab.NativeMethods]::SendModifiedKeyPress(
            [ExplorerCopyTab.NativeMethods]::VK_CONTROL,
            [ExplorerCopyTab.NativeMethods]::VK_C
        )
        if ($ctrlCResult -eq 0) {
            $ctrlCError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Log "Capture source Ctrl+C result=0 Win32Error=$ctrlCError"
        }
        else {
            Write-Log "Capture source Ctrl+C result=$ctrlCResult (expected 4)"
        }

        $capturedText = $null
        $clipboardWait = [System.Diagnostics.Stopwatch]::StartNew()
        while ($clipboardWait.ElapsedMilliseconds -lt 1000) {
            try {
                $capturedText = Get-Clipboard -Raw -Format Text
            }
            catch {
                $capturedText = $null
            }

            if (-not [string]::IsNullOrWhiteSpace($capturedText) -and $capturedText -ne $sentinel) {
                break
            }

            Start-Sleep -Milliseconds 50
        }

        [void][ExplorerCopyTab.NativeMethods]::SendKeyPress([ExplorerCopyTab.NativeMethods]::VK_ESCAPE)
        Start-Sleep -Milliseconds 50

        if (-not [string]::IsNullOrWhiteSpace($capturedText) -and $capturedText -ne $sentinel) {
            $normalized = $capturedText.Trim()
            Write-Log "Captured active tab target from UI: $normalized"
            return $normalized
        }

        Write-Log "Failed to capture active tab target from UI. ClipboardText=$capturedText Falling back to COM path=$($Context.Path)"
        return $Context.Path
    }
    finally {
        if ($clipboardHadText) {
            try {
                Set-Clipboard -Value $clipboardBackup
            }
            catch {
                Write-Log "Clipboard restore failed after capture: $_"
            }
        }
        else {
            try {
                Set-Clipboard -Value ""
            }
            catch {
                Write-Log "Clipboard clear failed after capture: $_"
            }
        }
    }
}

function Invoke-ExplorerTabClone {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Context
    )

    [ExplorerCopyTab.NativeMethods]::ReleaseModifierKeys()
    Start-Sleep -Milliseconds 30
    $foregroundSet = [ExplorerCopyTab.NativeMethods]::SetForegroundWindow($Context.Hwnd)
    Write-Log "Invoking clone for hwnd=$($Context.Hwnd.ToInt64()) path=$($Context.Path)"
    $foregroundReady = Wait-ForWindowForeground -Hwnd $Context.Hwnd
    Write-Log "Foreground request result=$foregroundSet ready=$foregroundReady"
    Start-Sleep -Milliseconds 180

    $navigationTarget = Get-ActiveExplorerNavigationTarget -Context $Context
    Write-Log "Using navigation target: $navigationTarget"

    $windowsBefore = @(Get-ExplorerWindowsForHwnd -Hwnd $Context.Hwnd)
    Write-Log "Explorer windows before Ctrl+T count=$($windowsBefore.Count)"

    Write-Log "Sending Ctrl+T"
    $ctrlTResult = [ExplorerCopyTab.NativeMethods]::SendModifiedKeyPress(
        [ExplorerCopyTab.NativeMethods]::VK_CONTROL,
        [ExplorerCopyTab.NativeMethods]::VK_T
    )
    if ($ctrlTResult -eq 0) {
        $ctrlTError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        Write-Log "Ctrl+T SendInput result=0 Win32Error=$ctrlTError"
    }
    else {
        Write-Log "Ctrl+T SendInput result=$ctrlTResult (expected 4)"
    }
    Start-Sleep -Milliseconds $NewTabDelayMs
    if ($PostNewTabSettleDelayMs -gt 0) {
        Start-Sleep -Milliseconds $PostNewTabSettleDelayMs
    }
    Write-Log "Post Ctrl+T wait complete delay=$NewTabDelayMs settle=$PostNewTabSettleDelayMs"

    $windowsAfter = @(Get-ExplorerWindowsForHwnd -Hwnd $Context.Hwnd)
    Write-Log "Explorer windows after Ctrl+T count=$($windowsAfter.Count)"

    $targetWindowInfo = $null
    if ($windowsAfter.Count -gt $windowsBefore.Count) {
        $targetWindowInfo = $windowsAfter[-1]
        Write-Log "Selected newest Explorer window entry after Ctrl+T title=$($targetWindowInfo.Title) path=$($targetWindowInfo.Path)"
    }
    else {
        $blankPathWindow = $windowsAfter | Where-Object { [string]::IsNullOrWhiteSpace($_.Path) } | Select-Object -First 1
        if ($null -ne $blankPathWindow) {
            $targetWindowInfo = $blankPathWindow
            Write-Log "Selected blank-path Explorer window entry title=$($targetWindowInfo.Title)"
        }
        elseif ($windowsAfter.Count -gt 0) {
            $targetWindowInfo = $windowsAfter[-1]
            Write-Log "Falling back to last Explorer window entry title=$($targetWindowInfo.Title) path=$($targetWindowInfo.Path)"
        }
    }

    if ($null -eq $targetWindowInfo) {
        Write-Log "No Explorer navigation target found after Ctrl+T."
    }
    else {
        try {
            $targetWindowInfo.Window.Navigate2($navigationTarget)
            Write-Log "Navigate2 invoked for path=$navigationTarget"
        }
        catch {
            Write-Log "Navigate2 failed: $_"
        }
    }

    Start-Sleep -Milliseconds $PostEnterVerifyDelayMs

    $verifyContext = Get-ActiveExplorerContext
    if ($null -eq $verifyContext) {
        Write-Log "Post-enter verification: no active Explorer context."
    }
    else {
        Write-Log "Post-enter verification: hwnd=$($verifyContext.Hwnd.ToInt64()) path=$($verifyContext.Path)"
    }
}

function Wait-ForHotKeyRelease {
    param(
        [int[]]$Keys = @(0x11, 0x12, 0x10),
        [int]$TimeoutMs = 1500
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($stopwatch.ElapsedMilliseconds -lt $TimeoutMs) {
        $pressed = $false
        foreach ($key in $keys) {
            if (([ExplorerCopyTab.NativeMethods]::GetAsyncKeyState($key) -band 0x8000) -ne 0) {
                $pressed = $true
                break
            }
        }

        if (-not $pressed) {
            return $true
        }

        Start-Sleep -Milliseconds 25
    }

    return $false
}

if ($RunOnce) {
    Write-Log "RunOnce started. Waiting up to $RunOnceWaitMs ms for Explorer foreground."
    $context = Wait-ForActiveExplorerContext -TimeoutMs $RunOnceWaitMs
    if ($null -eq $context) {
        Write-Log "RunOnce failed. No Explorer foreground within timeout."
        throw "No supported Explorer window became foreground within $RunOnceWaitMs ms."
    }

    Invoke-ExplorerTabClone -Context $context
    Write-Log "RunOnce completed."
    return
}

$resolvedHotKey = Resolve-HotKeyDefinition -Value $HotKey
$hotKeyId = 0x4554

if (-not [ExplorerCopyTab.NativeMethods]::RegisterHotKey([IntPtr]::Zero, $hotKeyId, $resolvedHotKey.Modifiers, $resolvedHotKey.VirtualKey)) {
    $errorCode = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
    Write-Log "RegisterHotKey failed. Win32Error=$errorCode HotKey=$HotKey"
    throw "RegisterHotKey failed with Win32 error code $errorCode. The shortcut may already be in use."
}

Write-Log "Listener started. HotKey=$HotKey LogPath=$LogPath"
Write-Host "Explorer copy-tab listener started. HotKey: $HotKey"
Write-Host "Press Ctrl+C in this window to stop."

$script:lastInvocationAt = [datetime]::MinValue

try {
    while ($true) {
        $messageId = [ExplorerCopyTab.NativeMethods]::WaitForNextHotKeyId()
        if ($messageId -eq -1) {
            throw "GetMessage failed while waiting for WM_HOTKEY."
        }

        if ($messageId -ne $hotKeyId) {
            continue
        }

        $now = Get-Date
        if (($now - $script:lastInvocationAt).TotalMilliseconds -lt 700) {
            Write-Log "HotKey ignored due to debounce."
            continue
        }
        $script:lastInvocationAt = $now

        Write-Log "HotKey received. Resolving active Explorer context."
        $context = Get-ActiveExplorerContext
        if ($null -eq $context) {
            Write-Log "HotKey received but active window was not a supported Explorer."
            continue
        }

        if (-not (Wait-ForHotKeyRelease -Keys @(0x11, 0x12, 0x10, [int]$resolvedHotKey.VirtualKey))) {
            Write-Log "HotKey release wait timed out. Proceeding anyway."
        }
        else {
            Write-Log "HotKey keys released. Starting input injection."
        }

        Invoke-ExplorerTabClone -Context $context
        Write-Log "HotKey processing completed."
    }
}
finally {
    [void][ExplorerCopyTab.NativeMethods]::UnregisterHotKey([IntPtr]::Zero, $hotKeyId)
    Write-Log "Listener stopped."
}
