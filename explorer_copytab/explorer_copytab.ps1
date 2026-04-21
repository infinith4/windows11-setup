param(
    [string]$HotKey = "Ctrl+Alt+Shift+D",
    [int]$NewTabDelayMs = 450,
    [int]$AddressBarDelayMs = 150,
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
            public KEYBDINPUT ki;
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

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern sbyte GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

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
                sbyte result = GetMessage(out message, IntPtr.Zero, 0, 0);
                if (result == 0)
                {
                    return 0;
                }

                if (result == -1)
                {
                    return -1;
                }

                if (message.message == WM_HOTKEY)
                {
                    return unchecked((int)message.wParam.ToUInt32());
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
                Window = $window
            }
        }
        catch {
            continue
        }
    }

    return $null
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

function Invoke-ExplorerTabClone {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Context
    )

    if ($null -ne $Context.Window) {
        try {
            Write-Log "Trying COM Navigate2 new-tab for path=$($Context.Path)"
            $Context.Window.Navigate2($Context.Path, 65536)
            Write-Log "COM Navigate2 accepted navOpenNewForegroundTab."
            return
        }
        catch {
            Write-Log "COM Navigate2 navOpenNewForegroundTab failed: $($_.Exception.Message)"
        }

        try {
            Write-Log "Trying COM Navigate2 open-in-new-tab for path=$($Context.Path)"
            $Context.Window.Navigate2($Context.Path, 2048)
            Write-Log "COM Navigate2 accepted navOpenInNewTab."
            return
        }
        catch {
            Write-Log "COM Navigate2 navOpenInNewTab failed: $($_.Exception.Message)"
        }
    }

    [void][ExplorerCopyTab.NativeMethods]::SetForegroundWindow($Context.Hwnd)
    Write-Log "Invoking SendInput fallback for hwnd=$($Context.Hwnd.ToInt64()) path=$($Context.Path)"
    Start-Sleep -Milliseconds 120

    [void][ExplorerCopyTab.NativeMethods]::SendModifiedKeyPress(
        [ExplorerCopyTab.NativeMethods]::VK_CONTROL,
        [ExplorerCopyTab.NativeMethods]::VK_T
    )
    Start-Sleep -Milliseconds $NewTabDelayMs

    [void][ExplorerCopyTab.NativeMethods]::SendModifiedKeyPress(
        [ExplorerCopyTab.NativeMethods]::VK_CONTROL,
        [ExplorerCopyTab.NativeMethods]::VK_L
    )
    Start-Sleep -Milliseconds $AddressBarDelayMs

    $clipboardBackup = $null
    $clipboardHadText = $false
    try {
        try {
            $clipboardBackup = Get-Clipboard -Raw -Format Text
            $clipboardHadText = $true
        }
        catch {
            $clipboardBackup = $null
            $clipboardHadText = $false
        }

        Set-Clipboard -Value $Context.Path
        Start-Sleep -Milliseconds 50
        [void][ExplorerCopyTab.NativeMethods]::SendModifiedKeyPress(
            [ExplorerCopyTab.NativeMethods]::VK_CONTROL,
            [ExplorerCopyTab.NativeMethods]::VK_A
        )
        Start-Sleep -Milliseconds 30
        [void][ExplorerCopyTab.NativeMethods]::SendModifiedKeyPress(
            [ExplorerCopyTab.NativeMethods]::VK_CONTROL,
            0x56
        )
        Start-Sleep -Milliseconds 50
    }
    finally {
        if ($clipboardHadText) {
            Set-Clipboard -Value $clipboardBackup
        }
    }

    [void][ExplorerCopyTab.NativeMethods]::SendKeyPress([ExplorerCopyTab.NativeMethods]::VK_RETURN)
}

function Wait-ForHotKeyRelease {
    param(
        [int]$TimeoutMs = 1500
    )

    $keys = @(0x11, 0x12, 0x10, 0x44)
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

        if (-not (Wait-ForHotKeyRelease)) {
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
