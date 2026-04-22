param(
    [ValidateSet("Listen", "Capture")]
    [string]$Mode = "Listen",

    [ValidateSet("Prompt", "FullScreen", "Region")]
    [string]$CaptureMode = "Prompt",

    [string]$OutputDirectory = "",
    [string]$LogPath = ""
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $PSScriptRoot "captures"
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $PSScriptRoot "screen_capture.log"
}

function Ensure-Directory {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        [void](New-Item -ItemType Directory -Path $Path -Force)
    }
}

Ensure-Directory -Path $OutputDirectory
$logDirectory = Split-Path -Path $LogPath -Parent
if (-not [string]::IsNullOrWhiteSpace($logDirectory)) {
    Ensure-Directory -Path $logDirectory
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    Add-Content -LiteralPath $LogPath -Value "$timestamp $Message"
}

if (-not ("ScreenCapture.NativeMethods" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ScreenCapture
{
    public static class NativeMethods
    {
        public const int WM_HOTKEY = 0x0312;
        public const uint PM_REMOVE = 0x0001;
        public const uint MOD_ALT = 0x0001;
        public const uint MOD_SHIFT = 0x0004;
        public const ushort VK_S = 0x53;

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }

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

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern sbyte GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport("user32.dll")]
        public static extern bool PeekMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);
    }

    public sealed class RegionSelectorForm : Form
    {
        private readonly Rectangle _virtualScreen;
        private Point _startPoint;
        private Point _currentPoint;
        private bool _dragging;
        private Rectangle _selectedRegion;

        public RegionSelectorForm()
        {
            _virtualScreen = SystemInformation.VirtualScreen;
            _selectedRegion = Rectangle.Empty;

            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            Bounds = _virtualScreen;
            TopMost = true;
            ShowInTaskbar = false;
            BackColor = Color.Black;
            Opacity = 0.25d;
            Cursor = Cursors.Cross;
            KeyPreview = true;
            DoubleBuffered = true;
        }

        public Rectangle SelectedRegion
        {
            get { return _selectedRegion; }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            Activate();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
                Close();
                return;
            }

            base.OnKeyDown(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);

            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            _startPoint = ToScreenPoint(e.Location);
            _currentPoint = _startPoint;
            _dragging = true;
            Invalidate();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);

            if (!_dragging)
            {
                return;
            }

            _currentPoint = ToScreenPoint(e.Location);
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);

            if (!_dragging)
            {
                return;
            }

            _currentPoint = ToScreenPoint(e.Location);
            _dragging = false;

            Rectangle normalized = Normalize(_startPoint, _currentPoint);
            if (normalized.Width >= 8 && normalized.Height >= 8)
            {
                _selectedRegion = normalized;
                DialogResult = DialogResult.OK;
            }
            else
            {
                DialogResult = DialogResult.Cancel;
            }

            Close();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            if (!_dragging)
            {
                return;
            }

            Rectangle selection = Normalize(_startPoint, _currentPoint);
            if (selection.Width <= 0 || selection.Height <= 0)
            {
                return;
            }

            Rectangle displayRectangle = new Rectangle(
                selection.X - _virtualScreen.X,
                selection.Y - _virtualScreen.Y,
                selection.Width,
                selection.Height
            );

            using (Brush fillBrush = new SolidBrush(Color.FromArgb(60, 0, 120, 215)))
            using (Pen outlinePen = new Pen(Color.FromArgb(220, 0, 120, 215), 2))
            {
                e.Graphics.FillRectangle(fillBrush, displayRectangle);
                e.Graphics.DrawRectangle(outlinePen, displayRectangle);
            }
        }

        private Point ToScreenPoint(Point localPoint)
        {
            return new Point(_virtualScreen.X + localPoint.X, _virtualScreen.Y + localPoint.Y);
        }

        private static Rectangle Normalize(Point a, Point b)
        {
            int left = Math.Min(a.X, b.X);
            int top = Math.Min(a.Y, b.Y);
            int right = Math.Max(a.X, b.X);
            int bottom = Math.Max(a.Y, b.Y);
            return Rectangle.FromLTRB(left, top, right, bottom);
        }

        public static Rectangle SelectRegion()
        {
            using (RegionSelectorForm form = new RegionSelectorForm())
            {
                DialogResult result = form.ShowDialog();
                return result == DialogResult.OK ? form.SelectedRegion : Rectangle.Empty;
            }
        }
    }
}
"@ -ReferencedAssemblies @("System.dll", "System.Drawing.dll", "System.Windows.Forms.dll")
}

function New-CaptureFilePath {
    param(
        [Parameter(Mandatory)]
        [string]$CaptureKind
    )

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
    $fileName = "capture_{0}_{1}.png" -f $timestamp, $CaptureKind.ToLowerInvariant()
    return Join-Path $OutputDirectory $fileName
}

function Get-NormalizedRectangle {
    param(
        [Parameter(Mandatory)]
        [System.Drawing.Point]$Start,

        [Parameter(Mandatory)]
        [System.Drawing.Point]$End
    )

    $left = [Math]::Min($Start.X, $End.X)
    $top = [Math]::Min($Start.Y, $End.Y)
    $right = [Math]::Max($Start.X, $End.X)
    $bottom = [Math]::Max($Start.Y, $End.Y)

    return [System.Drawing.Rectangle]::FromLTRB($left, $top, $right, $bottom)
}

function Save-ScreenArea {
    param(
        [Parameter(Mandatory)]
        [System.Drawing.Rectangle]$Bounds,

        [Parameter(Mandatory)]
        [string]$CaptureKind
    )

    if ($Bounds.Width -le 0 -or $Bounds.Height -le 0) {
        throw "Capture bounds must be larger than zero."
    }

    $filePath = New-CaptureFilePath -CaptureKind $CaptureKind
    $bitmap = New-Object System.Drawing.Bitmap($Bounds.Width, $Bounds.Height)

    try {
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        try {
            $sourcePoint = New-Object System.Drawing.Point($Bounds.X, $Bounds.Y)
            $targetPoint = [System.Drawing.Point]::Empty
            try {
                $graphics.CopyFromScreen($sourcePoint, $targetPoint, $Bounds.Size)
            }
            catch {
                throw "Screen copy failed. Run the script in an interactive desktop session and avoid protected windows. Details: $($_.Exception.Message)"
            }

            $bitmap.Save($filePath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally {
            $graphics.Dispose()
        }
    }
    finally {
        $bitmap.Dispose()
    }

    Write-Log "Capture saved. kind=$CaptureKind path=$filePath bounds=$($Bounds.X),$($Bounds.Y),$($Bounds.Width),$($Bounds.Height)"
    Write-Host "Saved: $filePath"
    return $filePath
}

function Show-CaptureModeSelector {
    $script:selection = $null
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Screen Capture"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(340, 170)
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true
    $form.KeyPreview = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Choose a capture mode.`r`nOpen this dialog with Alt+Shift+S."
    $label.AutoSize = $false
    $label.Size = New-Object System.Drawing.Size(300, 45)
    $label.Location = New-Object System.Drawing.Point(20, 15)
    $label.TextAlign = "MiddleCenter"

    $fullButton = New-Object System.Windows.Forms.Button
    $fullButton.Text = "Full Screen (&F)"
    $fullButton.Size = New-Object System.Drawing.Size(90, 34)
    $fullButton.Location = New-Object System.Drawing.Point(25, 80)

    $regionButton = New-Object System.Windows.Forms.Button
    $regionButton.Text = "Select Region (&R)"
    $regionButton.Size = New-Object System.Drawing.Size(90, 34)
    $regionButton.Location = New-Object System.Drawing.Point(123, 80)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Size = New-Object System.Drawing.Size(90, 34)
    $cancelButton.Location = New-Object System.Drawing.Point(221, 80)

    $fullButton.Add_Click({
        $script:selection = "FullScreen"
        $form.Close()
    })

    $regionButton.Add_Click({
        $script:selection = "Region"
        $form.Close()
    })

    $cancelButton.Add_Click({
        $script:selection = $null
        $form.Close()
    })

    $form.Add_KeyDown({
        param($sender, $eventArgs)

        switch ($eventArgs.KeyCode) {
            "F" {
                $script:selection = "FullScreen"
                $form.Close()
            }
            "R" {
                $script:selection = "Region"
                $form.Close()
            }
            "Escape" {
                $script:selection = $null
                $form.Close()
            }
        }
    })

    $form.Controls.AddRange(@($label, $fullButton, $regionButton, $cancelButton))
    $form.AcceptButton = $fullButton
    $form.CancelButton = $cancelButton

    [void]$form.ShowDialog()
    $form.Dispose()

    return $script:selection
}

function Select-ScreenRegion {
    $selection = [ScreenCapture.RegionSelectorForm]::SelectRegion()
    if ($selection.Width -le 0 -or $selection.Height -le 0) {
        return $null
    }

    Write-Log "Region selected. bounds=$($selection.X),$($selection.Y),$($selection.Width),$($selection.Height)"
    return $selection
}

function Get-CaptureBounds {
    param(
        [Parameter(Mandatory)]
        [ValidateSet("FullScreen", "Region")]
        [string]$Selection
    )

    switch ($Selection) {
        "FullScreen" {
            return [System.Windows.Forms.SystemInformation]::VirtualScreen
        }
        "Region" {
            return (Select-ScreenRegion)
        }
    }
}

function Invoke-ImageCapture {
    param(
        [Parameter(Mandatory)]
        [ValidateSet("Prompt", "FullScreen", "Region")]
        [string]$Selection
    )

    $resolvedSelection = $Selection
    if ($resolvedSelection -eq "Prompt") {
        $resolvedSelection = Show-CaptureModeSelector
    }

    if ([string]::IsNullOrWhiteSpace($resolvedSelection)) {
        Write-Log "Capture canceled at mode selection."
        return $null
    }

    Write-Log "Capture requested. selection=$resolvedSelection"
    Start-Sleep -Milliseconds 180

    $bounds = Get-CaptureBounds -Selection $resolvedSelection
    if ($null -eq $bounds) {
        Write-Log "Capture canceled before bounds were resolved. selection=$resolvedSelection"
        return $null
    }

    Start-Sleep -Milliseconds 120
    return Save-ScreenArea -Bounds $bounds -CaptureKind $resolvedSelection
}

function Wait-ForHotKeyRelease {
    param(
        [int]$TimeoutMs = 1500
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($stopwatch.ElapsedMilliseconds -lt $TimeoutMs) {
        $altPressed = ([ScreenCapture.NativeMethods]::GetAsyncKeyState(0x12) -band 0x8000) -ne 0
        $shiftPressed = ([ScreenCapture.NativeMethods]::GetAsyncKeyState(0x10) -band 0x8000) -ne 0
        $sPressed = ([ScreenCapture.NativeMethods]::GetAsyncKeyState(0x53) -band 0x8000) -ne 0

        if (-not ($altPressed -or $shiftPressed -or $sPressed)) {
            return $true
        }

        Start-Sleep -Milliseconds 25
    }

    return $false
}

function Clear-PendingHotKeyMessages {
    $message = New-Object ScreenCapture.NativeMethods+MSG
    while ([ScreenCapture.NativeMethods]::PeekMessage([ref]$message, [IntPtr]::Zero, [ScreenCapture.NativeMethods]::WM_HOTKEY, [ScreenCapture.NativeMethods]::WM_HOTKEY, [ScreenCapture.NativeMethods]::PM_REMOVE)) {
        Start-Sleep -Milliseconds 1
    }
}

function Start-HotKeyListener {
    $hotKeyId = 0x5343
    $modifiers = [ScreenCapture.NativeMethods]::MOD_ALT -bor [ScreenCapture.NativeMethods]::MOD_SHIFT
    $virtualKey = [ScreenCapture.NativeMethods]::VK_S

    if (-not [ScreenCapture.NativeMethods]::RegisterHotKey([IntPtr]::Zero, $hotKeyId, $modifiers, $virtualKey)) {
        $errorCode = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw "RegisterHotKey failed with Win32 error code $errorCode."
    }

    Write-Log "Listener started. HotKey=Alt+Shift+S OutputDirectory=$OutputDirectory"
    Write-Host "Screen capture listener started. HotKey: Alt+Shift+S"
    Write-Host "Press Ctrl+C in this window to stop."

    $lastInvocationAt = [datetime]::MinValue

    try {
        while ($true) {
            $message = New-Object ScreenCapture.NativeMethods+MSG
            $result = [ScreenCapture.NativeMethods]::GetMessage([ref]$message, [IntPtr]::Zero, 0, 0)
            if ($result -eq -1) {
                throw "GetMessage failed while waiting for WM_HOTKEY."
            }

            if ($result -eq 0) {
                break
            }

            if ($message.message -ne [ScreenCapture.NativeMethods]::WM_HOTKEY) {
                continue
            }

            if ($message.wParam.ToUInt32() -ne $hotKeyId) {
                continue
            }

            $now = Get-Date
            if (($now - $lastInvocationAt).TotalMilliseconds -lt 700) {
                Write-Log "HotKey ignored due to debounce."
                continue
            }

            $lastInvocationAt = $now
            [void](Wait-ForHotKeyRelease)
            Clear-PendingHotKeyMessages

            try {
                [void](Invoke-ImageCapture -Selection "Prompt")
            }
            catch {
                Write-Log "Capture failed after hotkey. $_"
                Write-Error $_
            }
        }
    }
    finally {
        [void][ScreenCapture.NativeMethods]::UnregisterHotKey([IntPtr]::Zero, $hotKeyId)
        Write-Log "Listener stopped."
    }
}

Write-Log "Script started. Mode=$Mode CaptureMode=$CaptureMode"

switch ($Mode) {
    "Capture" {
        [void](Invoke-ImageCapture -Selection $CaptureMode)
    }
    "Listen" {
        Start-HotKeyListener
    }
}
