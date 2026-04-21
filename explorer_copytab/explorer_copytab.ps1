param(
    [string]$HotKey = "Alt+Shift+Z",
    [int]$PostHotKeyCooldownMs = 500,
    [int]$NewTabDelayMs = 450,
    [int]$AddressBarDelayMs = 150,
    [int]$PostNewTabSettleDelayMs = 350,
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
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace ExplorerCopyTab
{
    [ComImport]
    [Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IServiceProvider
    {
        [PreserveSig]
        int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObject);
    }

    [ComImport]
    [Guid("00000114-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleWindow
    {
        void GetWindow(out IntPtr phwnd);
        void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
    }

    [ComImport]
    [Guid("000214E2-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellBrowser
    {
        void GetWindow(out IntPtr phwnd);
        void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
        void InsertMenusSB(IntPtr hmenuShared, IntPtr lpMenuWidths);
        void SetMenuSB(IntPtr hmenuShared, IntPtr holemenuRes, IntPtr hwndActiveObject);
        void RemoveMenusSB(IntPtr hmenuShared);
        void SetStatusTextSB([MarshalAs(UnmanagedType.LPWStr)] string pszStatusText);
        void EnableModelessSB([MarshalAs(UnmanagedType.Bool)] bool fEnable);
        void TranslateAcceleratorSB(IntPtr pmsg, ushort wID);
        void BrowseObject(IntPtr pidl, uint wFlags);
        void GetViewStateStream(uint grfMode, out IntPtr ppStrm);
        void GetControlWindow(uint id, out IntPtr phwnd);
        void SendControlMsg(uint id, uint uMsg, uint wParam, uint lParam, out IntPtr pret);
        void QueryActiveShellView(out IShellView ppshv);
        void OnViewWindowActive(IShellView ppshv);
        void SetToolbarItems(IntPtr lpButtons, uint nButtons, uint uFlags);
    }

    [ComImport]
    [Guid("000214E3-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellView
    {
        void GetWindow(out IntPtr phwnd);
        void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
        void TranslateAccelerator(IntPtr pmsg);
        void EnableModeless([MarshalAs(UnmanagedType.Bool)] bool fEnable);
        void UIActivate(uint uState);
        void Refresh();
        void CreateViewWindow(IShellView psvPrevious, IntPtr pfs, IShellBrowser psb, IntPtr prcView, out IntPtr phWnd);
        void DestroyViewWindow();
        void GetCurrentInfo(IntPtr pfs);
        void AddPropertySheetPages(uint dwReserved, IntPtr pfn, IntPtr lparam);
        void SaveViewState();
        void SelectItem(IntPtr pidlItem, uint uFlags);
        void GetItemObject(uint uItem, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppv);
    }

    public static class NativeMethods
    {
        public const int WM_HOTKEY = 0x0312;
        public const int WM_COMMAND = 0x0111;
        public const int EXPLORER_COMMAND_NEW_TAB = 0xA21B;
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

        private static readonly Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");
        private static readonly Guid IID_IShellBrowser = new Guid("000214E2-0000-0000-C000-000000000046");
        private static readonly Guid IID_IOleWindow = new Guid("00000114-0000-0000-C000-000000000046");
        private static readonly Guid SID_STopLevelBrowser = new Guid("4C96BE40-915C-11CF-99D3-00AA004AE837");
        private static readonly Guid SID_SShellBrowser = new Guid("000214E2-0000-0000-C000-000000000046");

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

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        public static string GetWindowClassName(IntPtr hWnd)
        {
            var builder = new StringBuilder(256);
            int length = GetClassName(hWnd, builder, builder.Capacity);
            return length > 0 ? builder.ToString() : string.Empty;
        }

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }

        private static T QueryService<T>(object browserWindow, Guid serviceId, Guid interfaceId) where T : class
        {
            var serviceProvider = browserWindow as IServiceProvider;
            if (serviceProvider == null)
            {
                return null;
            }

            IntPtr interfacePointer = IntPtr.Zero;
            try
            {
                int hr = serviceProvider.QueryService(ref serviceId, ref interfaceId, out interfacePointer);
                if (hr != 0 || interfacePointer == IntPtr.Zero)
                {
                    return null;
                }

                return Marshal.GetObjectForIUnknown(interfacePointer) as T;
            }
            finally
            {
                if (interfacePointer != IntPtr.Zero)
                {
                    Marshal.Release(interfacePointer);
                }
            }
        }

        private static IShellBrowser GetShellBrowser(object browserWindow)
        {
            return QueryService<IShellBrowser>(browserWindow, SID_STopLevelBrowser, IID_IShellBrowser)
                ?? QueryService<IShellBrowser>(browserWindow, SID_SShellBrowser, IID_IShellBrowser);
        }

        private static IOleWindow GetShellOleWindow(object browserWindow)
        {
            return QueryService<IOleWindow>(browserWindow, SID_STopLevelBrowser, IID_IOleWindow)
                ?? QueryService<IOleWindow>(browserWindow, SID_SShellBrowser, IID_IOleWindow);
        }

        private static string GetPathFromShellFolderViewDispatch(object backgroundDispatch)
        {
            if (backgroundDispatch == null)
            {
                return string.Empty;
            }

            object folder = null;
            object self = null;

            try
            {
                folder = backgroundDispatch.GetType().InvokeMember(
                    "Folder",
                    BindingFlags.GetProperty,
                    null,
                    backgroundDispatch,
                    null);

                if (folder == null)
                {
                    return string.Empty;
                }

                self = folder.GetType().InvokeMember(
                    "Self",
                    BindingFlags.GetProperty,
                    null,
                    folder,
                    null);

                if (self == null)
                {
                    return string.Empty;
                }

                object path = self.GetType().InvokeMember(
                    "Path",
                    BindingFlags.GetProperty,
                    null,
                    self,
                    null);

                return Convert.ToString(path) ?? string.Empty;
            }
            finally
            {
                ReleaseComObject(self);
                ReleaseComObject(folder);
            }
        }

        public static string GetActiveShellPath(object browserWindow)
        {
            IShellBrowser shellBrowser = null;
            IShellView shellView = null;
            object backgroundDispatch = null;

            try
            {
                shellBrowser = GetShellBrowser(browserWindow);
                if (shellBrowser == null)
                {
                    return string.Empty;
                }

                shellBrowser.QueryActiveShellView(out shellView);
                if (shellView == null)
                {
                    return string.Empty;
                }

                Guid dispatchId = IID_IDispatch;
                shellView.GetItemObject(0, ref dispatchId, out backgroundDispatch);
                return GetPathFromShellFolderViewDispatch(backgroundDispatch);
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComObject(backgroundDispatch);
                ReleaseComObject(shellView);
                ReleaseComObject(shellBrowser);
            }
        }

        public static IntPtr GetShellTabWindowHandle(object browserWindow)
        {
            IOleWindow oleWindow = null;

            try
            {
                oleWindow = GetShellOleWindow(browserWindow);
                if (oleWindow == null)
                {
                    return IntPtr.Zero;
                }

                IntPtr hwnd;
                oleWindow.GetWindow(out hwnd);
                return hwnd;
            }
            catch
            {
                return IntPtr.Zero;
            }
            finally
            {
                ReleaseComObject(oleWindow);
            }
        }

        public static IntPtr SendOpenNewTabCommand(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero)
            {
                return IntPtr.Zero;
            }

            return SendMessage(hWnd, WM_COMMAND, new IntPtr(EXPLORER_COMMAND_NEW_TAB), IntPtr.Zero);
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

function Initialize-UIAutomation {
    if ($script:UIAutomationInitialized) {
        return $script:UIAutomationAvailable
    }

    $script:UIAutomationInitialized = $true
    $script:UIAutomationAvailable = $false

    try {
        Add-Type -AssemblyName UIAutomationClient
        Add-Type -AssemblyName UIAutomationTypes
        $script:UIAutomationAvailable = $true
        Write-Log "UI Automation assemblies loaded."
    }
    catch {
        Write-Log "UI Automation assemblies failed to load: $_"
    }

    return $script:UIAutomationAvailable
}

function Get-SelectedExplorerTabTitle {
    param(
        [Parameter(Mandatory)]
        [IntPtr]$ShellTabHwnd
    )

    if ($ShellTabHwnd -eq [IntPtr]::Zero) {
        return $null
    }

    if (-not (Initialize-UIAutomation)) {
        return $null
    }

    try {
        $rootCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NativeWindowHandleProperty,
            [int]$ShellTabHwnd.ToInt64()
        )
        $root = [System.Windows.Automation.AutomationElement]::RootElement.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            $rootCondition
        )
        if ($null -eq $root) {
            Write-Log "Selected tab lookup: UIA root not found for shellTabHwnd=$($ShellTabHwnd.ToInt64())"
            return $null
        }

        $tabItemCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::TabItem
        )
        $tabItems = $root.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            $tabItemCondition
        )

        for ($i = 0; $i -lt $tabItems.Count; $i++) {
            $tabItem = $tabItems.Item($i)
            $isSelected = $false

            try {
                $selectionPattern = [System.Windows.Automation.SelectionItemPattern]$tabItem.GetCurrentPattern(
                    [System.Windows.Automation.SelectionItemPattern]::Pattern
                )
                $isSelected = $selectionPattern.Current.IsSelected
            }
            catch {
                try {
                    $isSelected = [bool]$tabItem.GetCurrentPropertyValue(
                        [System.Windows.Automation.SelectionItemPattern]::IsSelectedProperty
                    )
                }
                catch {
                    $isSelected = $false
                }
            }

            if (-not $isSelected) {
                continue
            }

            $name = [string]$tabItem.Current.Name
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $selectedTitle = $name.Trim()
                Write-Log "Selected tab lookup succeeded. shellTabHwnd=$($ShellTabHwnd.ToInt64()) title=$selectedTitle"
                return $selectedTitle
            }
        }

        Write-Log "Selected tab lookup found no selected TabItem for shellTabHwnd=$($ShellTabHwnd.ToInt64())"
    }
    catch {
        Write-Log "Selected tab lookup failed for shellTabHwnd=$($ShellTabHwnd.ToInt64()): $_"
    }

    return $null
}

function Get-FocusedExplorerElementInfo {
    if (-not (Initialize-UIAutomation)) {
        return $null
    }

    try {
        $focusedElement = [System.Windows.Automation.AutomationElement]::FocusedElement
        if ($null -eq $focusedElement) {
            Write-Log "Focused element lookup returned null."
            return $null
        }

        $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
        $ancestorNames = @()
        $current = $focusedElement
        for ($depth = 0; $depth -lt 12 -and $null -ne $current; $depth++) {
            $name = [string]$current.Current.Name
            $controlType = [string]$current.Current.ControlType.ProgrammaticName
            $automationId = [string]$current.Current.AutomationId
            $ancestorNames += "{0}:{1}:{2}" -f $controlType, $automationId, $name
            $current = $walker.GetParent($current)
        }

        $snapshot = [pscustomobject]@{
            Name = [string]$focusedElement.Current.Name
            ControlType = [string]$focusedElement.Current.ControlType.ProgrammaticName
            AutomationId = [string]$focusedElement.Current.AutomationId
            Ancestors = $ancestorNames
        }

        Write-Log "Focused element lookup succeeded. name=$($snapshot.Name) controlType=$($snapshot.ControlType) automationId=$($snapshot.AutomationId) ancestors=$($ancestorNames -join ' > ')"
        return $snapshot
    }
    catch {
        Write-Log "Focused element lookup failed: $_"
        return $null
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
    $candidates = @()

    foreach ($window in @($shell.Windows())) {
        try {
            $windowHandle = [IntPtr]::new([int64]$window.HWND)
            if ($windowHandle -ne $hwnd) {
                continue
            }

            $path = $null
            if ([string]::IsNullOrWhiteSpace($path)) {
                $path = $window.Document.Folder.Self.Path
            }

            if ([string]::IsNullOrWhiteSpace($path)) {
                continue
            }

            $shellTabHwnd = [ExplorerCopyTab.NativeMethods]::GetShellTabWindowHandle($window)
            $focusedItemName = $null
            try {
                $focusedItem = $window.Document.FocusedItem
                if ($null -ne $focusedItem) {
                    $focusedItemName = [string]$focusedItem.Name
                }
            }
            catch {
                $focusedItemName = $null
            }

            $candidates += [pscustomobject]@{
                Hwnd = $hwnd
                Path = $path
                ShellTabHwnd = $shellTabHwnd
                Title = [string]$window.LocationName
                FocusedItemName = $focusedItemName
                Window = $window
            }
        }
        catch {
            continue
        }
    }

    if ($candidates.Count -eq 0) {
        Write-Log "No Explorer candidates matched foreground hwnd=$($hwnd.ToInt64())"
        return $null
    }

    $shellTabHwnd = $candidates |
        Where-Object { $_.ShellTabHwnd -ne [IntPtr]::Zero } |
        Select-Object -First 1 -ExpandProperty ShellTabHwnd
    if ($null -eq $shellTabHwnd) {
        $shellTabHwnd = [IntPtr]::Zero
    }

    $shellTabClass = [ExplorerCopyTab.NativeMethods]::GetWindowClassName($shellTabHwnd)
    $selectedTitle = Get-SelectedExplorerTabTitle -ShellTabHwnd $shellTabHwnd
    $focusedElement = Get-FocusedExplorerElementInfo
    $activeShellPath = $null

    if (-not [string]::IsNullOrWhiteSpace($selectedTitle)) {
        $titleMatches = @($candidates | Where-Object { $_.Title -eq $selectedTitle })
        if ($titleMatches.Count -gt 0) {
            $selectedCandidate = $titleMatches[0]
            Write-Log "Matched Explorer via selected tab title. hwnd=$($hwnd.ToInt64()) path=$($selectedCandidate.Path) shellTabHwnd=$($shellTabHwnd.ToInt64()) shellTabClass=$shellTabClass title=$($selectedCandidate.Title)"
            return $selectedCandidate
        }

        Write-Log "Selected tab title did not match any candidate title. shellTabHwnd=$($shellTabHwnd.ToInt64()) selectedTitle=$selectedTitle candidates=$($candidates.Count)"
    }

    if ($null -ne $focusedElement -and -not [string]::IsNullOrWhiteSpace($focusedElement.Name)) {
        $focusedItemMatches = @(
            $candidates |
                Where-Object {
                    -not [string]::IsNullOrWhiteSpace($_.FocusedItemName) -and
                    $_.FocusedItemName -eq $focusedElement.Name
                }
        )
        if ($focusedItemMatches.Count -eq 1) {
            $selectedCandidate = $focusedItemMatches[0]
            Write-Log "Matched Explorer via focused item name. hwnd=$($hwnd.ToInt64()) path=$($selectedCandidate.Path) shellTabHwnd=$($shellTabHwnd.ToInt64()) shellTabClass=$shellTabClass title=$($selectedCandidate.Title) focusedItem=$($selectedCandidate.FocusedItemName)"
            return $selectedCandidate
        }

        if ($focusedItemMatches.Count -gt 1) {
            Write-Log "Focused item name matched multiple candidates. name=$($focusedElement.Name) count=$($focusedItemMatches.Count)"
        }
    }

    if ($null -ne $focusedElement -and $focusedElement.Ancestors.Count -gt 0) {
        foreach ($ancestorEntry in $focusedElement.Ancestors) {
            $parts = $ancestorEntry -split ":", 3
            if ($parts.Count -lt 3) {
                continue
            }

            $ancestorName = $parts[2]
            if ([string]::IsNullOrWhiteSpace($ancestorName)) {
                continue
            }

            $titleMatches = @($candidates | Where-Object { $_.Title -eq $ancestorName })
            if ($titleMatches.Count -eq 1) {
                $selectedCandidate = $titleMatches[0]
                Write-Log "Matched Explorer via focused ancestor title. hwnd=$($hwnd.ToInt64()) path=$($selectedCandidate.Path) shellTabHwnd=$($shellTabHwnd.ToInt64()) shellTabClass=$shellTabClass title=$($selectedCandidate.Title) ancestorName=$ancestorName"
                return $selectedCandidate
            }
        }
    }

    $activeShellPath = [ExplorerCopyTab.NativeMethods]::GetActiveShellPath($candidates[0].Window)
    if (-not [string]::IsNullOrWhiteSpace($activeShellPath)) {
        $pathMatches = @($candidates | Where-Object { $_.Path -eq $activeShellPath })
        if ($pathMatches.Count -gt 0) {
            $selectedCandidate = $pathMatches[0]
            Write-Log "Matched Explorer via active shell path fallback. hwnd=$($hwnd.ToInt64()) path=$($selectedCandidate.Path) shellTabHwnd=$($shellTabHwnd.ToInt64()) shellTabClass=$shellTabClass title=$($selectedCandidate.Title)"
            return $selectedCandidate
        }
    }

    $selectedCandidate = $candidates[0]
    Write-Log "Matched Explorer via first candidate fallback. hwnd=$($hwnd.ToInt64()) path=$($selectedCandidate.Path) shellTabHwnd=$($shellTabHwnd.ToInt64()) shellTabClass=$shellTabClass title=$($selectedCandidate.Title)"
    return $selectedCandidate
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

function Invoke-ExplorerTabClone {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Context
    )

    $foregroundSet = [ExplorerCopyTab.NativeMethods]::SetForegroundWindow($Context.Hwnd)
    Write-Log "Invoking clone for hwnd=$($Context.Hwnd.ToInt64()) path=$($Context.Path)"
    $foregroundReady = Wait-ForWindowForeground -Hwnd $Context.Hwnd
    Write-Log "Foreground request result=$foregroundSet ready=$foregroundReady"
    Start-Sleep -Milliseconds 180

    $navigationTarget = $Context.Path
    Write-Log "Using navigation target: $navigationTarget"

    if ($Context.ShellTabHwnd -eq [IntPtr]::Zero) {
        Write-Log "No ShellTabWindow handle was available for the active Explorer."
        return
    }

    $windowsBefore = @(Get-ExplorerWindowsForHwnd -Hwnd $Context.Hwnd)
    Write-Log "Explorer windows before new-tab command count=$($windowsBefore.Count)"

    Write-Log "Sending Explorer new-tab command to shellTabHwnd=$($Context.ShellTabHwnd.ToInt64())"
    $newTabCommandResult = [ExplorerCopyTab.NativeMethods]::SendOpenNewTabCommand($Context.ShellTabHwnd)
    Write-Log "New-tab command result=$($newTabCommandResult.ToInt64())"

    Start-Sleep -Milliseconds $NewTabDelayMs
    if ($PostNewTabSettleDelayMs -gt 0) {
        Start-Sleep -Milliseconds $PostNewTabSettleDelayMs
    }
    Write-Log "Post new-tab command wait complete delay=$NewTabDelayMs settle=$PostNewTabSettleDelayMs"

    $windowsAfter = @(Get-ExplorerWindowsForHwnd -Hwnd $Context.Hwnd)
    Write-Log "Explorer windows after new-tab command count=$($windowsAfter.Count)"

    $targetWindowInfo = $null
    if ($windowsAfter.Count -gt $windowsBefore.Count) {
        $targetWindowInfo = Compare-Object -ReferenceObject $windowsBefore -DifferenceObject $windowsAfter -Property Path,Title -PassThru |
            Where-Object { $_.SideIndicator -eq "=>" } |
            Select-Object -First 1
        if ($null -eq $targetWindowInfo) {
            $targetWindowInfo = $windowsAfter[-1]
        }
        Write-Log "Selected newest Explorer window entry after new-tab command title=$($targetWindowInfo.Title) path=$($targetWindowInfo.Path)"
    }
    else {
        $blankPathWindow = $windowsAfter | Where-Object { [string]::IsNullOrWhiteSpace($_.Path) } | Select-Object -First 1
        if ($null -ne $blankPathWindow) {
            $targetWindowInfo = $blankPathWindow
            Write-Log "Selected blank-path Explorer window entry title=$($targetWindowInfo.Title)"
        }
    }

    if ($null -eq $targetWindowInfo) {
        Write-Log "No Explorer navigation target found after new-tab command."
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

        $hotKeyReleased = Wait-ForHotKeyRelease -Keys @(0x11, 0x12, 0x10, [int]$resolvedHotKey.VirtualKey)
        if (-not $hotKeyReleased) {
            Write-Log "HotKey release wait timed out. Proceeding anyway."
            [ExplorerCopyTab.NativeMethods]::ReleaseModifierKeys()
            Write-Log "Forced modifier release after timeout."
            Start-Sleep -Milliseconds 30
        }
        else {
            Write-Log "HotKey keys released."
        }

        if ($PostHotKeyCooldownMs -gt 0) {
            Write-Log "Post-hotkey cooldown delay=$PostHotKeyCooldownMs"
            Start-Sleep -Milliseconds $PostHotKeyCooldownMs
        }

        Invoke-ExplorerTabClone -Context $context
        Write-Log "HotKey processing completed."
    }
}
finally {
    [void][ExplorerCopyTab.NativeMethods]::UnregisterHotKey([IntPtr]::Zero, $hotKeyId)
    Write-Log "Listener stopped."
}
