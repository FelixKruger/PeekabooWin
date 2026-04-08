param(
    [Parameter(Mandatory = $true)]
    [string]$RequestJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$script:WinRtAsTaskMethod = $null

if (-not ("PeekabooWindows.NativeMethods" -as [type])) {
    Add-Type -ReferencedAssemblies @("System.Drawing") -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace PeekabooWindows {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT {
        public int type;
        public InputUnion U;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct InputUnion {
        [FieldOffset(0)]
        public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct KEYBDINPUT {
        public ushort wVk;
        public ushort wScan;
        public int dwFlags;
        public int time;
        public IntPtr dwExtraInfo;
    }

    public static class NativeMethods {
        public const int INPUT_KEYBOARD = 1;
        public const int KEYEVENTF_KEYUP = 0x0002;
        public const int KEYEVENTF_UNICODE = 0x0004;
        public const uint SW_MAXIMIZE = 3;
        public const uint SW_MINIMIZE = 6;
        public const uint SW_RESTORE = 9;
        public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        public const uint MOUSEEVENTF_LEFTUP = 0x0004;
        public const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        public const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        public const uint MOUSEEVENTF_WHEEL = 0x0800;

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int maxCount);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, uint cmdShow);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdc, uint flags);

        [DllImport("user32.dll")]
        public static extern bool MoveWindow(IntPtr hWnd, int x, int y, int width, int height, bool repaint);

        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();

        [DllImport("user32.dll")]
        public static extern bool SetProcessDpiAwarenessContext(IntPtr value);

        [DllImport("user32.dll")]
        public static extern uint GetDpiForWindow(IntPtr hWnd);

        public static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);

        [DllImport("user32.dll")]
        public static extern uint SendInput(uint count, INPUT[] inputs, int size);

        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extraInfo);

        public static string GetWindowTitle(IntPtr hWnd) {
            var sb = new StringBuilder(1024);
            GetWindowText(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public static string GetWindowClassName(IntPtr hWnd) {
            var sb = new StringBuilder(512);
            GetClassName(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public static Rectangle GetBounds(IntPtr hWnd) {
            RECT rect;
            if (!GetWindowRect(hWnd, out rect)) {
                return Rectangle.Empty;
            }

            return Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom);
        }

        public static List<object> ListWindows() {
            var windows = new List<object>();

            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam) {
                if (!IsWindowVisible(hWnd)) {
                    return true;
                }

                string title = GetWindowTitle(hWnd);
                if (string.IsNullOrWhiteSpace(title)) {
                    return true;
                }

                uint processId;
                GetWindowThreadProcessId(hWnd, out processId);
                Rectangle bounds = GetBounds(hWnd);
                string processName = "";

                try {
                    using (var process = Process.GetProcessById((int)processId)) {
                        processName = process.ProcessName;
                    }
                } catch {
                }

                windows.Add(new {
                    hwnd = "0x" + hWnd.ToInt64().ToString("X"),
                    title = title,
                    className = GetWindowClassName(hWnd),
                    processId = processId,
                    processName = processName,
                    bounds = new {
                        left = bounds.Left,
                        top = bounds.Top,
                        width = bounds.Width,
                        height = bounds.Height
                    }
                });

                return true;
            }, IntPtr.Zero);

            return windows;
        }

        public static string FindWindowHandleByTitle(string titleSubstring) {
            IntPtr matched = IntPtr.Zero;

            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam) {
                if (!IsWindowVisible(hWnd)) {
                    return true;
                }

                string title = GetWindowTitle(hWnd);
                if (string.IsNullOrWhiteSpace(title)) {
                    return true;
                }

                if (title.IndexOf(titleSubstring, StringComparison.OrdinalIgnoreCase) >= 0) {
                    matched = hWnd;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            if (matched == IntPtr.Zero) {
                return null;
            }

            return "0x" + matched.ToInt64().ToString("X");
        }

        public static void MoveCursor(int x, int y) {
            SetCursorPos(x, y);
        }

        public static void MouseClick(string button, bool doubleClick) {
            uint down = button == "right" ? MOUSEEVENTF_RIGHTDOWN : MOUSEEVENTF_LEFTDOWN;
            uint up = button == "right" ? MOUSEEVENTF_RIGHTUP : MOUSEEVENTF_LEFTUP;
            int count = doubleClick ? 2 : 1;

            for (int index = 0; index < count; index++) {
                mouse_event(down, 0, 0, 0, UIntPtr.Zero);
                mouse_event(up, 0, 0, 0, UIntPtr.Zero);
                System.Threading.Thread.Sleep(50);
            }
        }

        public static void MouseButtonDown(string button) {
            uint down = button == "right" ? MOUSEEVENTF_RIGHTDOWN : MOUSEEVENTF_LEFTDOWN;
            mouse_event(down, 0, 0, 0, UIntPtr.Zero);
        }

        public static void MouseButtonUp(string button) {
            uint up = button == "right" ? MOUSEEVENTF_RIGHTUP : MOUSEEVENTF_LEFTUP;
            mouse_event(up, 0, 0, 0, UIntPtr.Zero);
        }

        public static void ScrollWheel(int delta, int repeats) {
            int count = Math.Max(1, repeats);

            for (int index = 0; index < count; index++) {
                mouse_event(MOUSEEVENTF_WHEEL, 0, 0, unchecked((uint)delta), UIntPtr.Zero);
                System.Threading.Thread.Sleep(35);
            }
        }

        public static void TypeText(string text) {
            foreach (char ch in text) {
                INPUT press = new INPUT();
                press.type = INPUT_KEYBOARD;
                press.U.ki = new KEYBDINPUT {
                    wScan = ch,
                    dwFlags = KEYEVENTF_UNICODE
                };

                INPUT release = new INPUT();
                release.type = INPUT_KEYBOARD;
                release.U.ki = new KEYBDINPUT {
                    wScan = ch,
                    dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP
                };

                INPUT[] inputs = new INPUT[] { press, release };
                SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
            }
        }
    }
}
"@
}

function Initialize-DpiAwareness {
    try {
        [void][PeekabooWindows.NativeMethods]::SetProcessDpiAwarenessContext([PeekabooWindows.NativeMethods]::DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
        return
    }
    catch {
    }

    try {
        [void][PeekabooWindows.NativeMethods]::SetProcessDPIAware()
    }
    catch {
    }
}

Initialize-DpiAwareness

function Write-JsonResult {
    param(
        [bool]$Ok,
        [object]$Result = $null,
        [string]$ErrorMessage = $null
    )

    $payload = if ($Ok) {
        @{
            ok     = $true
            result = $Result
        }
    }
    else {
        @{
            ok    = $false
            error = $ErrorMessage
        }
    }

    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $payload | ConvertTo-Json -Depth 8 -Compress
}

function Get-VisibleAppEntries {
    $windows = @([PeekabooWindows.NativeMethods]::ListWindows())
    $groups = $windows | Group-Object processId
    $apps = @()

    foreach ($group in $groups) {
        $primary = $group.Group | Select-Object -First 1
        $titles = @($group.Group | ForEach-Object { [string]$_.title } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        $apps += ,([pscustomobject]@{
                processId   = [int]$primary.processId
                processName = [string]$primary.processName
                windowCount = $group.Count
                titles      = $titles
                hwnds       = @($group.Group | ForEach-Object { [string]$_.hwnd })
            })
    }

    return @($apps | Sort-Object processName, processId)
}

function Get-ScreenEntries {
    $screens = [System.Windows.Forms.Screen]::AllScreens
    $results = @()

    for ($index = 0; $index -lt $screens.Length; $index++) {
        $screen = $screens[$index]
        $bounds = $screen.Bounds
        $workingArea = $screen.WorkingArea

        $results += ,([pscustomobject]@{
                index       = $index
                deviceName  = [string]$screen.DeviceName
                isPrimary   = [bool]$screen.Primary
                bounds      = @{
                    left   = $bounds.Left
                    top    = $bounds.Top
                    width  = $bounds.Width
                    height = $bounds.Height
                }
                workingArea = @{
                    left   = $workingArea.Left
                    top    = $workingArea.Top
                    width  = $workingArea.Width
                    height = $workingArea.Height
                }
            })
    }

    return $results
}

function Resolve-CaptureBounds {
    param(
        [object]$Payload
    )

    $screenIndexProperty = $Payload.PSObject.Properties["screenIndex"]
    if ($null -ne $screenIndexProperty -and -not [string]::IsNullOrWhiteSpace([string]$screenIndexProperty.Value)) {
        $screenIndex = [int]$screenIndexProperty.Value
        $screens = @(Get-ScreenEntries)

        if ($screenIndex -lt 0 -or $screenIndex -ge $screens.Count) {
            throw "Invalid screen index '$screenIndex'"
        }

        return $screens[$screenIndex].bounds
    }

    $virtualScreen = [System.Windows.Forms.SystemInformation]::VirtualScreen
    return @{
        left   = $virtualScreen.Left
        top    = $virtualScreen.Top
        width  = $virtualScreen.Width
        height = $virtualScreen.Height
    }
}

function Resolve-AppProcess {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $processIdProperty = $Payload.PSObject.Properties["processId"]
    if ($null -ne $processIdProperty -and -not [string]::IsNullOrWhiteSpace([string]$processIdProperty.Value)) {
        return Get-Process -Id ([int]$processIdProperty.Value) -ErrorAction Stop
    }

    $nameProperty = $Payload.PSObject.Properties["name"]
    $titleProperty = $Payload.PSObject.Properties["title"]
    $matchModeProperty = $Payload.PSObject.Properties["matchMode"]
    $matchMode = if ($null -ne $matchModeProperty -and -not [string]::IsNullOrWhiteSpace([string]$matchModeProperty.Value)) {
        [string]$matchModeProperty.Value
    }
    else {
        "contains"
    }

    $apps = @(Get-VisibleAppEntries)

    if ($null -ne $nameProperty -and -not [string]::IsNullOrWhiteSpace([string]$nameProperty.Value)) {
        $name = [string]$nameProperty.Value
        $matches = @($apps | Where-Object {
                $candidate = [string]$_.processName
                if ($matchMode -eq "exact") {
                    $candidate -ceq $name
                }
                else {
                    $candidate.IndexOf($name, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
                }
            })

        if ($matches.Count -lt 1) {
            throw "No visible app matched name '$name'"
        }

        if ($matches.Count -gt 1) {
            throw "App selector '$name' matched $($matches.Count) visible apps; use --process-id or a more specific selector"
        }

        return Get-Process -Id ([int]$matches[0].processId) -ErrorAction Stop
    }

    if ($null -ne $titleProperty -and -not [string]::IsNullOrWhiteSpace([string]$titleProperty.Value)) {
        $title = [string]$titleProperty.Value
        $windowMatch = @([PeekabooWindows.NativeMethods]::ListWindows() | Where-Object {
                $candidate = [string]$_.title
                if ($matchMode -eq "exact") {
                    $candidate -ceq $title
                }
                else {
                    $candidate.IndexOf($title, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
                }
            } | Select-Object -First 1)

        if ($windowMatch.Count -lt 1) {
            throw "No visible app window matched title '$title'"
        }

        return Get-Process -Id ([int]$windowMatch[0].processId) -ErrorAction Stop
    }

    throw "Missing processId, name, or title"
}

function Resolve-AppWindowHandle {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $process = Resolve-AppProcess -Payload $Payload
    $windows = @([PeekabooWindows.NativeMethods]::ListWindows() | Where-Object { [int]$_.processId -eq [int]$process.Id })

    if ($windows.Count -lt 1) {
        throw "No visible window found for process '$($process.ProcessName)'"
    }

    return ConvertTo-Hwnd -Value ([string]$windows[0].hwnd)
}

function Resolve-LaunchedAppResult {
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.Process]$Process,
        [Parameter(Mandatory = $true)]
        [string]$Command
    )

    Start-Sleep -Milliseconds 700

    $visibleApps = @(Get-VisibleAppEntries)
    $fallbackName = [System.IO.Path]::GetFileNameWithoutExtension($Command)
    $processId = [int]$Process.Id
    $processName = $fallbackName
    $launchExited = $false

    try {
        $Process.Refresh()
        $launchExited = $Process.HasExited
        if (-not $launchExited -and -not [string]::IsNullOrWhiteSpace([string]$Process.ProcessName)) {
            $processName = [string]$Process.ProcessName
        }
    }
    catch {
        $launchExited = $true
    }

    $matchedApp = @($visibleApps | Where-Object { [int]$_.processId -eq $processId } | Select-Object -First 1)

    if ($matchedApp.Count -lt 1) {
        $candidateNames = @()

        if (-not [string]::IsNullOrWhiteSpace($processName)) {
            $candidateNames += $processName
        }

        if (-not [string]::IsNullOrWhiteSpace($fallbackName) -and -not ($candidateNames -contains $fallbackName)) {
            $candidateNames += $fallbackName
        }

        foreach ($candidateName in $candidateNames) {
            $nameMatches = @($visibleApps | Where-Object { [string]$_.processName -ieq $candidateName })

            if ($nameMatches.Count -eq 1) {
                $matchedApp = @($nameMatches[0])
                break
            }
        }
    }

    if ($matchedApp.Count -gt 0) {
        return @{
            processId      = [int]$matchedApp[0].processId
            processName    = [string]$matchedApp[0].processName
            titles         = @($matchedApp[0].titles)
            hwnds          = @($matchedApp[0].hwnds)
            resolved       = $true
            launchExited   = [bool]$launchExited
        }
    }

    return @{
        processId    = $processId
        processName  = $processName
        titles       = @()
        hwnds        = @()
        resolved     = $false
        launchExited = [bool]$launchExited
    }
}

function ConvertTo-Hwnd {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Value
    )

    $text = [string]$Value
    if ($text.StartsWith("0x")) {
        return [IntPtr]([Convert]::ToInt64($text.Substring(2), 16))
    }

    return [IntPtr]([Convert]::ToInt64($text, 10))
}

function Resolve-WindowHandle {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $hwndProperty = $Payload.PSObject.Properties["hwnd"]
    if ($null -ne $hwndProperty -and -not [string]::IsNullOrWhiteSpace([string]$hwndProperty.Value)) {
        return ConvertTo-Hwnd -Value $hwndProperty.Value
    }

    $titleProperty = $Payload.PSObject.Properties["title"]
    if ($null -eq $titleProperty -or [string]::IsNullOrWhiteSpace([string]$titleProperty.Value)) {
        throw "Missing hwnd or title"
    }

    $resolvedHandle = [PeekabooWindows.NativeMethods]::FindWindowHandleByTitle([string]$titleProperty.Value)

    if ([string]::IsNullOrWhiteSpace($resolvedHandle)) {
        throw "No matching window found for title '$($titleProperty.Value)'"
    }

    return ConvertTo-Hwnd -Value $resolvedHandle
}

function Save-ScreenCapture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [object]$Bounds = $null
    )

    $captureBounds = if ($null -ne $Bounds) {
        $Bounds
    }
    else {
        Resolve-CaptureBounds -Payload ([pscustomobject]@{})
    }

    $bitmap = New-Object System.Drawing.Bitmap $captureBounds.width, $captureBounds.height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

    try {
        $graphics.CopyFromScreen($captureBounds.left, $captureBounds.top, 0, 0, $bitmap.Size)
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }

    return @{
        path   = $Path
        bounds = @{
            left   = [int]$captureBounds.left
            top    = [int]$captureBounds.top
            width  = [int]$captureBounds.width
            height = [int]$captureBounds.height
        }
    }
}

function Test-BlackFrame {
    param(
        [Parameter(Mandatory = $true)]
        [System.Drawing.Bitmap]$Bitmap,
        [int]$SampleCount = 20,
        [int]$BrightnessThreshold = 5
    )

    $width = $Bitmap.Width
    $height = $Bitmap.Height

    if ($width -le 0 -or $height -le 0) {
        return $true
    }

    $totalBrightness = 0
    $seed = 42
    $rng = New-Object System.Random $seed

    for ($i = 0; $i -lt $SampleCount; $i++) {
        $px = $rng.Next(0, $width)
        $py = $rng.Next(0, $height)
        $pixel = $Bitmap.GetPixel($px, $py)
        $totalBrightness += ([int]$pixel.R + [int]$pixel.G + [int]$pixel.B) / 3
    }

    $averageBrightness = $totalBrightness / $SampleCount
    return $averageBrightness -lt $BrightnessThreshold
}

function Save-WindowCapture {
    param(
        [Parameter(Mandatory = $true)]
        [IntPtr]$Handle,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $bounds = [PeekabooWindows.NativeMethods]::GetBounds($Handle)
    if ($bounds.Width -le 0 -or $bounds.Height -le 0) {
        throw "Window has invalid bounds"
    }

    $captureMethod = $null

    # Tier 1: PrintWindow with PW_RENDERFULLCONTENT (flag 2) for GPU/DWM-composited apps
    $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $hdc = $graphics.GetHdc()

    try {
        $printed = [PeekabooWindows.NativeMethods]::PrintWindow($Handle, $hdc, 2)
        $graphics.ReleaseHdc($hdc)
        $hdc = [IntPtr]::Zero

        if ($printed -and -not (Test-BlackFrame -Bitmap $bitmap)) {
            $captureMethod = "PrintWindow-FullContent"
        }
    }
    catch {
        if ($hdc -ne [IntPtr]::Zero) {
            $graphics.ReleaseHdc($hdc)
            $hdc = [IntPtr]::Zero
        }
    }
    finally {
        if ($null -eq $captureMethod) {
            $graphics.Dispose()
            $bitmap.Dispose()
        }
    }

    # Tier 2: PrintWindow with classic flag 0 for traditional GDI apps
    if ($null -eq $captureMethod) {
        $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $hdc = $graphics.GetHdc()

        try {
            $printed = [PeekabooWindows.NativeMethods]::PrintWindow($Handle, $hdc, 0)
            $graphics.ReleaseHdc($hdc)
            $hdc = [IntPtr]::Zero

            if ($printed -and -not (Test-BlackFrame -Bitmap $bitmap)) {
                $captureMethod = "PrintWindow-Classic"
            }
        }
        catch {
            if ($hdc -ne [IntPtr]::Zero) {
                $graphics.ReleaseHdc($hdc)
                $hdc = [IntPtr]::Zero
            }
        }
        finally {
            if ($null -eq $captureMethod) {
                $graphics.Dispose()
                $bitmap.Dispose()
            }
        }
    }

    # Tier 3: CopyFromScreen as final fallback (captures screen pixels at window position)
    if ($null -eq $captureMethod) {
        $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

        try {
            $graphics.CopyFromScreen($bounds.Left, $bounds.Top, 0, 0, $bitmap.Size)
            $captureMethod = "CopyFromScreen"
        }
        finally {
            $graphics.Dispose()
        }
    }

    try {
        $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $bitmap.Dispose()
    }

    return @{
        path          = $Path
        hwnd          = ("0x{0:X}" -f $Handle.ToInt64())
        captureMethod = $captureMethod
        bounds        = @{
            left   = $bounds.Left
            top    = $bounds.Top
            width  = $bounds.Width
            height = $bounds.Height
        }
    }
}

function Get-ControlTypeObject {
    param(
        [string]$ControlType
    )

    if ([string]::IsNullOrWhiteSpace($ControlType)) {
        return $null
    }

    $normalized = $ControlType.Trim().ToLowerInvariant()
    switch ($normalized) {
        "button" { return [System.Windows.Automation.ControlType]::Button }
        "edit" { return [System.Windows.Automation.ControlType]::Edit }
        "textbox" { return [System.Windows.Automation.ControlType]::Edit }
        "text" { return [System.Windows.Automation.ControlType]::Text }
        "menuitem" { return [System.Windows.Automation.ControlType]::MenuItem }
        "menu" { return [System.Windows.Automation.ControlType]::Menu }
        "checkbox" { return [System.Windows.Automation.ControlType]::CheckBox }
        "radiobutton" { return [System.Windows.Automation.ControlType]::RadioButton }
        "tab" { return [System.Windows.Automation.ControlType]::Tab }
        "tabitem" { return [System.Windows.Automation.ControlType]::TabItem }
        "list" { return [System.Windows.Automation.ControlType]::List }
        "listitem" { return [System.Windows.Automation.ControlType]::ListItem }
        "tree" { return [System.Windows.Automation.ControlType]::Tree }
        "treeitem" { return [System.Windows.Automation.ControlType]::TreeItem }
        "combobox" { return [System.Windows.Automation.ControlType]::ComboBox }
        "window" { return [System.Windows.Automation.ControlType]::Window }
        default { throw "Unsupported control type '$ControlType'" }
    }
}

function Get-FriendlyControlTypeName {
    param(
        [object]$ControlType
    )

    if ($null -eq $ControlType) {
        return ""
    }

    $programmaticName = [string]$ControlType.ProgrammaticName
    if ([string]::IsNullOrWhiteSpace($programmaticName)) {
        return ""
    }

    $lastDot = $programmaticName.LastIndexOf(".")
    if ($lastDot -ge 0 -and $lastDot -lt ($programmaticName.Length - 1)) {
        return $programmaticName.Substring($lastDot + 1).ToLowerInvariant()
    }

    return $programmaticName.ToLowerInvariant()
}

function Resolve-AutomationRoot {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Payload
    )

    $hwndProperty = $Payload.PSObject.Properties["hwnd"]
    $titleProperty = $Payload.PSObject.Properties["title"]

    if (($null -ne $hwndProperty -and -not [string]::IsNullOrWhiteSpace([string]$hwndProperty.Value)) -or
        ($null -ne $titleProperty -and -not [string]::IsNullOrWhiteSpace([string]$titleProperty.Value))) {
        $hwnd = Resolve-WindowHandle -Payload $Payload
        return [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
    }

    return [System.Windows.Automation.AutomationElement]::RootElement
}

function Test-RectangleVisible {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rect,
        [object]$VisibleBounds = $null
    )

    if ($null -eq $Rect) {
        return $false
    }

    $left = [double]$Rect.Left
    $top = [double]$Rect.Top
    $width = [double]$Rect.Width
    $height = [double]$Rect.Height

    if ([double]::IsNaN($left) -or [double]::IsNaN($top) -or [double]::IsNaN($width) -or [double]::IsNaN($height)) {
        return $false
    }

    if ($width -le 1 -or $height -le 1) {
        return $false
    }

    if ($null -eq $VisibleBounds) {
        return $true
    }

    $visibleLeft = [double]$VisibleBounds.left
    $visibleTop = [double]$VisibleBounds.top
    $visibleRight = $visibleLeft + [double]$VisibleBounds.width
    $visibleBottom = $visibleTop + [double]$VisibleBounds.height
    $right = $left + $width
    $bottom = $top + $height

    return $right -gt $visibleLeft -and $left -lt $visibleRight -and $bottom -gt $visibleTop -and $top -lt $visibleBottom
}

function Find-AutomationElements {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Root,
        [object]$Payload,
        [object]$VisibleBounds = $null,
        [int]$DefaultMaxResults = 10
    )

    $nameProperty = $Payload.PSObject.Properties["name"]
    $automationIdProperty = $Payload.PSObject.Properties["automationId"]
    $classNameProperty = $Payload.PSObject.Properties["className"]
    $controlTypeProperty = $Payload.PSObject.Properties["controlType"]
    $maxResultsProperty = $Payload.PSObject.Properties["maxResults"]
    $matchModeProperty = $Payload.PSObject.Properties["matchMode"]

    $nameFilter = if ($null -ne $nameProperty) { [string]$nameProperty.Value } else { "" }
    $automationIdFilter = if ($null -ne $automationIdProperty) { [string]$automationIdProperty.Value } else { "" }
    $classNameFilter = if ($null -ne $classNameProperty) { [string]$classNameProperty.Value } else { "" }
    $matchMode = if ($null -ne $matchModeProperty -and -not [string]::IsNullOrWhiteSpace([string]$matchModeProperty.Value)) {
        [string]$matchModeProperty.Value
    }
    else {
        "contains"
    }
    $maxResults = if ($null -ne $maxResultsProperty -and [int]$maxResultsProperty.Value -gt 0) {
        [int]$maxResultsProperty.Value
    }
    else {
        $DefaultMaxResults
    }

    $condition = if ($null -ne $controlTypeProperty -and -not [string]::IsNullOrWhiteSpace([string]$controlTypeProperty.Value)) {
        $controlType = Get-ControlTypeObject -ControlType ([string]$controlTypeProperty.Value)
        New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            $controlType
        )
    }
    else {
        [System.Windows.Automation.Condition]::TrueCondition
    }

    $collection = $Root.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
    $results = @()

    foreach ($element in $collection) {
        $currentName = [string]$element.Current.Name
        $currentAutomationId = [string]$element.Current.AutomationId
        $currentClassName = [string]$element.Current.ClassName
        $friendlyControlType = Get-FriendlyControlTypeName -ControlType $element.Current.ControlType
        $rect = $element.Current.BoundingRectangle

        if (-not (Test-RectangleVisible -Rect $rect -VisibleBounds $VisibleBounds)) {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($currentName) -and [string]::IsNullOrWhiteSpace($currentAutomationId)) {
            continue
        }

        if (-not [string]::IsNullOrWhiteSpace($nameFilter)) {
            if ($matchMode -eq "exact" -and $currentName -cne $nameFilter) {
                continue
            }

            if ($matchMode -ne "exact" -and $currentName.IndexOf($nameFilter, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                continue
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($automationIdFilter)) {
            if ($matchMode -eq "exact" -and $currentAutomationId -cne $automationIdFilter) {
                continue
            }

            if ($matchMode -ne "exact" -and $currentAutomationId.IndexOf($automationIdFilter, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                continue
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($classNameFilter)) {
            if ($matchMode -eq "exact" -and $currentClassName -cne $classNameFilter) {
                continue
            }

            if ($matchMode -ne "exact" -and $currentClassName.IndexOf($classNameFilter, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
                continue
            }
        }

        $results += ,@{
                name         = $currentName
                automationId = $currentAutomationId
                className    = $currentClassName
                controlType  = $friendlyControlType
                controlTypeRaw = [string]$element.Current.ControlType.ProgrammaticName
                isEnabled    = [bool]$element.Current.IsEnabled
                bounds       = @{
                    left   = [int][Math]::Round($rect.Left)
                    top    = [int][Math]::Round($rect.Top)
                    width  = [int][Math]::Round($rect.Width)
                    height = [int][Math]::Round($rect.Height)
                }
                center       = @{
                    x = [int][Math]::Round($rect.Left + ($rect.Width / 2))
                    y = [int][Math]::Round($rect.Top + ($rect.Height / 2))
                }
            }

        if ($results.Count -ge $maxResults) {
            break
        }
    }

    return $results
}

function Add-ElementIdentifiers {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Elements
    )

    $results = @()
    $index = 1

    foreach ($element in $Elements) {
        $record = [ordered]@{
            id = "e$index"
        }

        foreach ($property in $element.GetEnumerator()) {
            $record[$property.Key] = $property.Value
        }

        $results += ,([pscustomobject]$record)
        $index += 1
    }

    return $results
}

function Save-AnnotatedCapture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$AnnotatedPath,
        [Parameter(Mandatory = $true)]
        [object]$CaptureBounds,
        [Parameter(Mandatory = $true)]
        [array]$Elements
    )

    $bitmap = [System.Drawing.Bitmap]::FromFile($Path)
    try {
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(220, 255, 59, 48)), 3
        $fillBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(45, 255, 59, 48))
        $labelBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(230, 255, 59, 48))
        $textBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
        $font = New-Object System.Drawing.Font ("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

        try {
            foreach ($element in $Elements) {
                $localLeft = [int]([double]$element.bounds.left - [double]$CaptureBounds.left)
                $localTop = [int]([double]$element.bounds.top - [double]$CaptureBounds.top)
                $localWidth = [int][Math]::Round([double]$element.bounds.width)
                $localHeight = [int][Math]::Round([double]$element.bounds.height)

                if ($localWidth -le 1 -or $localHeight -le 1) {
                    continue
                }

                $rect = New-Object System.Drawing.Rectangle ($localLeft, $localTop, $localWidth, $localHeight)
                $graphics.FillRectangle($fillBrush, $rect)
                $graphics.DrawRectangle($pen, $rect)

                $labelText = if ([string]::IsNullOrWhiteSpace([string]$element.name)) {
                    [string]$element.id
                }
                else {
                    "$($element.id) $($element.name)"
                }

                if ($labelText.Length -gt 48) {
                    $labelText = $labelText.Substring(0, 45) + "..."
                }

                $labelSize = $graphics.MeasureString($labelText, $font)
                $labelX = [float][Math]::Max(0, $localLeft)
                $labelY = [float][Math]::Max(0, $localTop - [int][Math]::Ceiling($labelSize.Height) - 4)
                $labelRect = New-Object System.Drawing.RectangleF ($labelX, $labelY, [float]($labelSize.Width + 10), [float]($labelSize.Height + 4))
                $graphics.FillRectangle($labelBrush, $labelRect)
                $graphics.DrawString($labelText, $font, $textBrush, $labelX + 5, $labelY + 2)
            }

            $bitmap.Save($AnnotatedPath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally {
            $font.Dispose()
            $textBrush.Dispose()
            $labelBrush.Dispose()
            $fillBrush.Dispose()
            $pen.Dispose()
            $graphics.Dispose()
        }
    }
    finally {
        $bitmap.Dispose()
    }
}

function Initialize-OcrRuntime {
    $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
    $null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]
    $null = [Windows.Media.Ocr.OcrResult, Windows.Media.Ocr, ContentType = WindowsRuntime]
}

function Get-WinRtAsTaskMethod {
    if ($script:WinRtAsTaskMethod -eq $null) {
        $script:WinRtAsTaskMethod = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
            $_.Name -eq "AsTask" -and $_.IsGenericMethod -and $_.GetGenericArguments().Count -eq 1 -and $_.GetParameters().Count -eq 1
        } | Select-Object -First 1
    }

    if ($script:WinRtAsTaskMethod -eq $null) {
        throw "Unable to resolve the WinRT AsTask bridge"
    }

    return $script:WinRtAsTaskMethod
}

function Await-WinRtOperation {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Operation,
        [Parameter(Mandatory = $true)]
        [Type]$ResultType
    )

    $asTaskMethod = Get-WinRtAsTaskMethod
    $genericMethod = $asTaskMethod.MakeGenericMethod(@($ResultType))
    $task = $genericMethod.Invoke($null, @($Operation))
    return $task.GetAwaiter().GetResult()
}

function New-BoundsCenter {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Bounds
    )

    return [pscustomobject]@{
        x = [int][Math]::Round([double]$Bounds.left + ([double]$Bounds.width / 2))
        y = [int][Math]::Round([double]$Bounds.top + ([double]$Bounds.height / 2))
    }
}

function Convert-OcrRectToBounds {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Rect,
        [double]$ScaleFactor = 1,
        [object]$CaptureBounds = $null
    )

    $offsetLeft = 0
    $offsetTop = 0
    if ($null -ne $CaptureBounds) {
        if ($CaptureBounds.PSObject.Properties["left"]) {
            $offsetLeft = [double]$CaptureBounds.left
        }
        if ($CaptureBounds.PSObject.Properties["top"]) {
            $offsetTop = [double]$CaptureBounds.top
        }
    }

    $safeScale = if ($ScaleFactor -gt 0) { $ScaleFactor } else { 1 }
    $left = [int][Math]::Round(($Rect.X / $safeScale) + $offsetLeft)
    $top = [int][Math]::Round(($Rect.Y / $safeScale) + $offsetTop)
    $width = [int][Math]::Max(1, [Math]::Round($Rect.Width / $safeScale))
    $height = [int][Math]::Max(1, [Math]::Round($Rect.Height / $safeScale))

    return [pscustomobject]@{
        left = $left
        top = $top
        width = $width
        height = $height
    }
}

function Prepare-OcrImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Initialize-OcrRuntime

    $sourceImage = [System.Drawing.Image]::FromFile($Path)
    try {
        $sourceWidth = [int]$sourceImage.Width
        $sourceHeight = [int]$sourceImage.Height
    }
    finally {
        $sourceImage.Dispose()
    }

    $maxDimension = [double][Windows.Media.Ocr.OcrEngine]::MaxImageDimension
    $largestDimension = [double][Math]::Max($sourceWidth, $sourceHeight)

    if ($largestDimension -le $maxDimension) {
        return [pscustomobject]@{
            path = $Path
            tempPath = $null
            scaleFactor = 1
            usedScaledImage = $false
            sourceWidth = $sourceWidth
            sourceHeight = $sourceHeight
        }
    }

    $scaleFactor = $maxDimension / $largestDimension
    $scaledWidth = [int][Math]::Max(1, [Math]::Round($sourceWidth * $scaleFactor))
    $scaledHeight = [int][Math]::Max(1, [Math]::Round($sourceHeight * $scaleFactor))
    $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("peekaboo-ocr-" + [System.Guid]::NewGuid().ToString("N") + ".png")

    $sourceBitmap = [System.Drawing.Bitmap]::FromFile($Path)
    $scaledBitmap = New-Object System.Drawing.Bitmap ($scaledWidth, $scaledHeight)
    $graphics = [System.Drawing.Graphics]::FromImage($scaledBitmap)
    try {
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $graphics.DrawImage($sourceBitmap, 0, 0, $scaledWidth, $scaledHeight)
        $scaledBitmap.Save($tempPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $graphics.Dispose()
        $scaledBitmap.Dispose()
        $sourceBitmap.Dispose()
    }

    return [pscustomobject]@{
        path = $tempPath
        tempPath = $tempPath
        scaleFactor = $scaleFactor
        usedScaledImage = $true
        sourceWidth = $sourceWidth
        sourceHeight = $sourceHeight
    }
}

function Invoke-ImageOcr {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [object]$CaptureBounds = $null
    )

    Initialize-OcrRuntime

    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if ($null -eq $engine) {
        throw "Windows OCR engine is not available"
    }

    $preparedImage = Prepare-OcrImage -Path $Path
    $stream = $null

    try {
        $file = Await-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($preparedImage.path)) ([Windows.Storage.StorageFile])
        $stream = Await-WinRtOperation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
        $decoder = Await-WinRtOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Await-WinRtOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $ocrResult = Await-WinRtOperation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])

        $normalizedLines = @()
        $wordCount = 0

        foreach ($line in @($ocrResult.Lines)) {
            $words = @()

            foreach ($word in @($line.Words)) {
                $bounds = Convert-OcrRectToBounds -Rect $word.BoundingRect -ScaleFactor ([double]$preparedImage.scaleFactor) -CaptureBounds $CaptureBounds
                $words += ,([pscustomobject]@{
                        text = [string]$word.Text
                        bounds = $bounds
                        center = New-BoundsCenter -Bounds $bounds
                    })
                $wordCount += 1
            }

            $lineBounds = $null
            if ($words.Count -gt 0) {
                $left = [double]::PositiveInfinity
                $top = [double]::PositiveInfinity
                $right = [double]::NegativeInfinity
                $bottom = [double]::NegativeInfinity

                foreach ($wordRecord in $words) {
                    $wordLeft = [double]$wordRecord.bounds.left
                    $wordTop = [double]$wordRecord.bounds.top
                    $wordRight = [double]$wordRecord.bounds.left + [double]$wordRecord.bounds.width
                    $wordBottom = [double]$wordRecord.bounds.top + [double]$wordRecord.bounds.height

                    if ($wordLeft -lt $left) { $left = $wordLeft }
                    if ($wordTop -lt $top) { $top = $wordTop }
                    if ($wordRight -gt $right) { $right = $wordRight }
                    if ($wordBottom -gt $bottom) { $bottom = $wordBottom }
                }

                $lineBounds = [pscustomobject]@{
                    left = [int][Math]::Round($left)
                    top = [int][Math]::Round($top)
                    width = [int][Math]::Max(1, [Math]::Round($right - $left))
                    height = [int][Math]::Max(1, [Math]::Round($bottom - $top))
                }
            }

            $normalizedLines += ,([pscustomobject]@{
                    text = [string]$line.Text
                    bounds = $lineBounds
                    center = if ($null -ne $lineBounds) { New-BoundsCenter -Bounds $lineBounds } else { $null }
                    words = $words
                })
        }

        return @{
            available = $true
            text = [string]$ocrResult.Text
            lineCount = $normalizedLines.Count
            wordCount = $wordCount
            recognizerLanguage = [string]$engine.RecognizerLanguage.LanguageTag
            sourceWidth = [int]$preparedImage.sourceWidth
            sourceHeight = [int]$preparedImage.sourceHeight
            usedScaledImage = [bool]$preparedImage.usedScaledImage
            scaleFactor = [double]$preparedImage.scaleFactor
            lines = $normalizedLines
        }
    }
    finally {
        if ($null -ne $stream -and $stream -is [System.IDisposable]) {
            $stream.Dispose()
        }

        if ($preparedImage.tempPath) {
            Remove-Item -LiteralPath $preparedImage.tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-AutomationElementClick {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Automation.AutomationElement]$Element
    )

    $invokePattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$invokePattern)) {
        ([System.Windows.Automation.InvokePattern]$invokePattern).Invoke()
        return "invoke"
    }

    $selectionPattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$selectionPattern)) {
        ([System.Windows.Automation.SelectionItemPattern]$selectionPattern).Select()
        return "select"
    }

    $togglePattern = $null
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$togglePattern)) {
        ([System.Windows.Automation.TogglePattern]$togglePattern).Toggle()
        return "toggle"
    }

    $rect = $Element.Current.BoundingRectangle
    if ($rect.Width -le 0 -or $rect.Height -le 0) {
        throw "Element does not expose a clickable bounding rectangle"
    }

    $x = [int][Math]::Round($rect.Left + ($rect.Width / 2))
    $y = [int][Math]::Round($rect.Top + ($rect.Height / 2))
    [PeekabooWindows.NativeMethods]::MoveCursor($x, $y)
    [PeekabooWindows.NativeMethods]::MouseClick("left", $false)
    return "mouse"
}

function Get-ScrollDelta {
    param(
        [object]$Payload
    )

    $directionProperty = $Payload.PSObject.Properties["direction"]
    $direction = if ($null -ne $directionProperty -and -not [string]::IsNullOrWhiteSpace([string]$directionProperty.Value)) {
        [string]$directionProperty.Value
    }
    else {
        "down"
    }

    switch ($direction.ToLowerInvariant()) {
        "down" { return -120 }
        "up" { return 120 }
        default { throw "Unsupported scroll direction '$direction'" }
    }
}

function Set-WindowForeground {
    param(
        [Parameter(Mandatory = $true)]
        [IntPtr]$Handle,
        [int]$ProcessId = 0,
        [string]$WindowTitle = ""
    )

    [void][PeekabooWindows.NativeMethods]::ShowWindow($Handle, [PeekabooWindows.NativeMethods]::SW_RESTORE)

    if ([PeekabooWindows.NativeMethods]::SetForegroundWindow($Handle)) {
        return $true
    }

    try {
        $shell = New-Object -ComObject WScript.Shell
        $null = $shell.SendKeys('%')

        $activated = $false
        if ($ProcessId -gt 0) {
            $activated = [bool]$shell.AppActivate($ProcessId)
        }
        elseif (-not [string]::IsNullOrWhiteSpace($WindowTitle)) {
            $activated = [bool]$shell.AppActivate($WindowTitle)
        }

        Start-Sleep -Milliseconds 120

        if ([PeekabooWindows.NativeMethods]::SetForegroundWindow($Handle)) {
            return $true
        }

        return $activated
    }
    catch {
        return $false
    }
}

function Set-WindowBounds {
    param(
        [Parameter(Mandatory = $true)]
        [IntPtr]$Handle,
        [Parameter(Mandatory = $true)]
        [int]$X,
        [Parameter(Mandatory = $true)]
        [int]$Y,
        [Parameter(Mandatory = $true)]
        [int]$Width,
        [Parameter(Mandatory = $true)]
        [int]$Height
    )

    if ($Width -le 0 -or $Height -le 0) {
        throw "Window width and height must be greater than zero"
    }

    $dpiScale = 1.0
    try {
        $dpi = [PeekabooWindows.NativeMethods]::GetDpiForWindow($Handle)
        if ($dpi -gt 0) {
            $dpiScale = [double]$dpi / 96.0
        }
    }
    catch {
    }

    $apiX = [int][Math]::Round($X / $dpiScale)
    $apiY = [int][Math]::Round($Y / $dpiScale)
    $apiWidth = [int][Math]::Round($Width / $dpiScale)
    $apiHeight = [int][Math]::Round($Height / $dpiScale)

    [void][PeekabooWindows.NativeMethods]::ShowWindow($Handle, [PeekabooWindows.NativeMethods]::SW_RESTORE)
    $moved = [PeekabooWindows.NativeMethods]::MoveWindow($Handle, $apiX, $apiY, $apiWidth, $apiHeight, $true)

    if (-not $moved) {
        throw "Unable to update the requested window bounds"
    }

    $bounds = [PeekabooWindows.NativeMethods]::GetBounds($Handle)
    return @{
        hwnd            = ("0x{0:X}" -f $Handle.ToInt64())
        requestedBounds = @{
            left   = $X
            top    = $Y
            width  = $Width
            height = $Height
        }
        bounds          = @{
            left   = $bounds.Left
            top    = $bounds.Top
            width  = $bounds.Width
            height = $bounds.Height
        }
        delta           = @{
            left   = $bounds.Left - $X
            top    = $bounds.Top - $Y
            width  = $bounds.Width - $Width
            height = $bounds.Height - $Height
        }
        exact           = ($bounds.Left -eq $X -and $bounds.Top -eq $Y -and $bounds.Width -eq $Width -and $bounds.Height -eq $Height)
    }
}

function Set-WindowState {
    param(
        [Parameter(Mandatory = $true)]
        [IntPtr]$Handle,
        [Parameter(Mandatory = $true)]
        [string]$State
    )

    $normalizedState = $State.ToLowerInvariant()
    $showCode = switch ($normalizedState) {
        "restore" { [PeekabooWindows.NativeMethods]::SW_RESTORE }
        "maximize" { [PeekabooWindows.NativeMethods]::SW_MAXIMIZE }
        "minimize" { [PeekabooWindows.NativeMethods]::SW_MINIMIZE }
        default { throw "Unsupported window state '$State'" }
    }

    $changed = [PeekabooWindows.NativeMethods]::ShowWindow($Handle, $showCode)
    Start-Sleep -Milliseconds 120
    $bounds = [PeekabooWindows.NativeMethods]::GetBounds($Handle)

    return @{
        hwnd    = ("0x{0:X}" -f $Handle.ToInt64())
        state   = $normalizedState
        changed = [bool]$changed
        bounds  = @{
            left   = $bounds.Left
            top    = $bounds.Top
            width  = $bounds.Width
            height = $bounds.Height
        }
    }
}

function Invoke-MouseDrag {
    param(
        [Parameter(Mandatory = $true)]
        [int]$FromX,
        [Parameter(Mandatory = $true)]
        [int]$FromY,
        [Parameter(Mandatory = $true)]
        [int]$ToX,
        [Parameter(Mandatory = $true)]
        [int]$ToY,
        [string]$Button = "left",
        [int]$Steps = 16,
        [int]$DurationMs = 300
    )

    $stepCount = [Math]::Max(1, $Steps)
    $sleepMs = [Math]::Max(0, [int][Math]::Floor($DurationMs / $stepCount))

    [PeekabooWindows.NativeMethods]::MoveCursor($FromX, $FromY)
    Start-Sleep -Milliseconds 40
    [PeekabooWindows.NativeMethods]::MouseButtonDown($Button)

    try {
        for ($index = 1; $index -le $stepCount; $index++) {
            $progress = [double]$index / [double]$stepCount
            $currentX = [int][Math]::Round($FromX + (($ToX - $FromX) * $progress))
            $currentY = [int][Math]::Round($FromY + (($ToY - $FromY) * $progress))
            [PeekabooWindows.NativeMethods]::MoveCursor($currentX, $currentY)

            if ($sleepMs -gt 0) {
                Start-Sleep -Milliseconds $sleepMs
            }
        }
    }
    finally {
        [PeekabooWindows.NativeMethods]::MouseButtonUp($Button)
    }

    return @{
        from = @{
            x = $FromX
            y = $FromY
        }
        to   = @{
            x = $ToX
            y = $ToY
        }
        button = $Button
        steps = $stepCount
        durationMs = $DurationMs
    }
}

function Invoke-ClipboardTextEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $originalClipboard = $null
    $hadClipboardData = $false

    try {
        $originalClipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
        $hadClipboardData = $null -ne $originalClipboard
    }
    catch {
    }

    try {
        [System.Windows.Forms.Clipboard]::SetText($Text)
        Start-Sleep -Milliseconds 40
        [System.Windows.Forms.SendKeys]::SendWait("^v")
        Start-Sleep -Milliseconds 40
    }
    finally {
        try {
            if ($hadClipboardData) {
                [System.Windows.Forms.Clipboard]::SetDataObject($originalClipboard, $true)
            }
            else {
                [System.Windows.Forms.Clipboard]::Clear()
            }
        }
        catch {
        }
    }
}

try {
    $request = $RequestJson | ConvertFrom-Json
    $action = [string]$request.action
    $payload = $request.payload

    switch ($action) {
        "list-windows" {
            Write-JsonResult -Ok $true -Result @{
                windows = [PeekabooWindows.NativeMethods]::ListWindows()
            }
            break
        }
        "focus-window" {
            $hwnd = Resolve-WindowHandle -Payload $payload
            $focused = Set-WindowForeground -Handle $hwnd -WindowTitle ([PeekabooWindows.NativeMethods]::GetWindowTitle($hwnd))

            if (-not $focused) {
                throw "Unable to focus the requested window"
            }

            Write-JsonResult -Ok $true -Result @{
                hwnd = ("0x{0:X}" -f $hwnd.ToInt64())
            }
            break
        }
        "move-window" {
            $hwnd = Resolve-WindowHandle -Payload $payload
            $currentBounds = [PeekabooWindows.NativeMethods]::GetBounds($hwnd)
            $x = [int][Math]::Round([double]$payload.x)
            $y = [int][Math]::Round([double]$payload.y)
            $result = Set-WindowBounds -Handle $hwnd -X $x -Y $y -Width $currentBounds.Width -Height $currentBounds.Height
            Write-JsonResult -Ok $true -Result $result
            break
        }
        "resize-window" {
            $hwnd = Resolve-WindowHandle -Payload $payload
            $currentBounds = [PeekabooWindows.NativeMethods]::GetBounds($hwnd)
            $width = [int][Math]::Round([double]$payload.width)
            $height = [int][Math]::Round([double]$payload.height)
            $result = Set-WindowBounds -Handle $hwnd -X $currentBounds.Left -Y $currentBounds.Top -Width $width -Height $height
            Write-JsonResult -Ok $true -Result $result
            break
        }
        "set-window-bounds" {
            $hwnd = Resolve-WindowHandle -Payload $payload
            $x = [int][Math]::Round([double]$payload.x)
            $y = [int][Math]::Round([double]$payload.y)
            $width = [int][Math]::Round([double]$payload.width)
            $height = [int][Math]::Round([double]$payload.height)
            $result = Set-WindowBounds -Handle $hwnd -X $x -Y $y -Width $width -Height $height
            Write-JsonResult -Ok $true -Result $result
            break
        }
        "set-window-state" {
            $hwnd = Resolve-WindowHandle -Payload $payload
            $stateProperty = $payload.PSObject.Properties["state"]
            if ($null -eq $stateProperty -or [string]::IsNullOrWhiteSpace([string]$stateProperty.Value)) {
                throw "Missing window state"
            }

            $result = Set-WindowState -Handle $hwnd -State ([string]$stateProperty.Value)
            Write-JsonResult -Ok $true -Result $result
            break
        }
        "launch-app" {
            $args = @()
            if ($null -ne $payload.args) {
                $args = @($payload.args)
            }

            if ($args.Count -gt 0) {
                $process = Start-Process -FilePath $payload.command -ArgumentList $args -PassThru
            }
            else {
                $process = Start-Process -FilePath $payload.command -PassThru
            }

            $resolvedLaunch = Resolve-LaunchedAppResult -Process $process -Command ([string]$payload.command)
            Write-JsonResult -Ok $true -Result @{
                processId    = [int]$resolvedLaunch.processId
                processName  = [string]$resolvedLaunch.processName
                command      = $payload.command
                args         = $args
                titles       = @($resolvedLaunch.titles)
                hwnds        = @($resolvedLaunch.hwnds)
                resolved     = [bool]$resolvedLaunch.resolved
                launchExited = [bool]$resolvedLaunch.launchExited
            }
            break
        }
        "list-apps" {
            Write-JsonResult -Ok $true -Result @{
                apps = @(Get-VisibleAppEntries)
            }
            break
        }
        "list-screens" {
            Write-JsonResult -Ok $true -Result @{
                screens = @(Get-ScreenEntries)
            }
            break
        }
        "switch-app" {
            $hwnd = Resolve-AppWindowHandle -Payload $payload
            $process = Resolve-AppProcess -Payload $payload
            $focused = Set-WindowForeground -Handle $hwnd -ProcessId ([int]$process.Id) -WindowTitle ([PeekabooWindows.NativeMethods]::GetWindowTitle($hwnd))

            if (-not $focused) {
                throw "Unable to switch to the requested app"
            }

            Write-JsonResult -Ok $true -Result @{
                processId   = [int]$process.Id
                processName = [string]$process.ProcessName
                hwnd        = ("0x{0:X}" -f $hwnd.ToInt64())
            }
            break
        }
        "quit-app" {
            $process = Resolve-AppProcess -Payload $payload
            $closed = $process.CloseMainWindow()
            $exited = $process.WaitForExit(3000)

            Write-JsonResult -Ok $true -Result @{
                processId   = [int]$process.Id
                processName = [string]$process.ProcessName
                closeSent   = [bool]$closed
                exited      = [bool]$exited
            }
            break
        }
        "move-mouse" {
            $x = [int][Math]::Round([double]$payload.x)
            $y = [int][Math]::Round([double]$payload.y)
            [PeekabooWindows.NativeMethods]::MoveCursor($x, $y)
            Write-JsonResult -Ok $true -Result @{ x = $x; y = $y }
            break
        }
        "get-cursor-position" {
            $position = [System.Windows.Forms.Cursor]::Position
            Write-JsonResult -Ok $true -Result @{
                x = [int]$position.X
                y = [int]$position.Y
            }
            break
        }
        "click" {
            $x = [int][Math]::Round([double]$payload.x)
            $y = [int][Math]::Round([double]$payload.y)
            [PeekabooWindows.NativeMethods]::MoveCursor($x, $y)
            [PeekabooWindows.NativeMethods]::MouseClick([string]$payload.button, [bool]$payload.double)
            Write-JsonResult -Ok $true -Result @{
                x      = $x
                y      = $y
                button = [string]$payload.button
                double = [bool]$payload.double
            }
            break
        }
        "drag" {
            $fromX = [int][Math]::Round([double]$payload.fromX)
            $fromY = [int][Math]::Round([double]$payload.fromY)
            $toX = [int][Math]::Round([double]$payload.toX)
            $toY = [int][Math]::Round([double]$payload.toY)
            $button = if ($null -ne $payload.button -and -not [string]::IsNullOrWhiteSpace([string]$payload.button)) {
                [string]$payload.button
            }
            else {
                "left"
            }
            $steps = if ($null -ne $payload.steps -and [int]$payload.steps -gt 0) {
                [int]$payload.steps
            }
            else {
                16
            }
            $durationMs = if ($null -ne $payload.durationMs -and [int]$payload.durationMs -ge 0) {
                [int]$payload.durationMs
            }
            else {
                300
            }

            $result = Invoke-MouseDrag -FromX $fromX -FromY $fromY -ToX $toX -ToY $toY -Button $button -Steps $steps -DurationMs $durationMs
            Write-JsonResult -Ok $true -Result $result
            break
        }
        "scroll" {
            $xProperty = $payload.PSObject.Properties["x"]
            $yProperty = $payload.PSObject.Properties["y"]

            if ($null -ne $xProperty -and $null -ne $yProperty) {
                $x = [int][Math]::Round([double]$xProperty.Value)
                $y = [int][Math]::Round([double]$yProperty.Value)
                [PeekabooWindows.NativeMethods]::MoveCursor($x, $y)
            }
            else {
                $position = [System.Windows.Forms.Cursor]::Position
                $x = [int]$position.X
                $y = [int]$position.Y
            }

            $ticksProperty = $payload.PSObject.Properties["ticks"]
            $ticks = if ($null -ne $ticksProperty -and [int]$ticksProperty.Value -gt 0) {
                [int]$ticksProperty.Value
            }
            else {
                3
            }

            $directionProperty = $payload.PSObject.Properties["direction"]
            $direction = if ($null -ne $directionProperty -and -not [string]::IsNullOrWhiteSpace([string]$directionProperty.Value)) {
                [string]$directionProperty.Value
            }
            else {
                "down"
            }

            $delta = Get-ScrollDelta -Payload $payload
            [PeekabooWindows.NativeMethods]::ScrollWheel($delta, $ticks)

            Write-JsonResult -Ok $true -Result @{
                x         = $x
                y         = $y
                direction = $direction
                ticks     = $ticks
            }
            break
        }
        "type-text" {
            Invoke-ClipboardTextEntry -Text ([string]$payload.text)
            Write-JsonResult -Ok $true -Result @{
                typed = [string]$payload.text
            }
            break
        }
        "ocr-image" {
            $result = Invoke-ImageOcr -Path ([string]$payload.path) -CaptureBounds $payload.bounds
            Write-JsonResult -Ok $true -Result $result
            break
        }
        "press-keys" {
            [System.Windows.Forms.SendKeys]::SendWait([string]$payload.keys)
            Write-JsonResult -Ok $true -Result @{
                keys = [string]$payload.keys
            }
            break
        }
        "capture-screen" {
            $captureBounds = Resolve-CaptureBounds -Payload $payload
            $result = Save-ScreenCapture -Path ([string]$payload.path) -Bounds $captureBounds
            Write-JsonResult -Ok $true -Result $result
            break
        }
        "capture-window" {
            $hwnd = Resolve-WindowHandle -Payload $payload
            $result = Save-WindowCapture -Handle $hwnd -Path ([string]$payload.path)
            Write-JsonResult -Ok $true -Result $result
            break
        }
        "see-ui" {
            $modeProperty = $payload.PSObject.Properties["mode"]
            $hwndProperty = $payload.PSObject.Properties["hwnd"]
            $titleProperty = $payload.PSObject.Properties["title"]

            $mode = if ($null -ne $modeProperty -and -not [string]::IsNullOrWhiteSpace([string]$modeProperty.Value)) {
                [string]$modeProperty.Value
            }
            else {
                "screen"
            }

            $root = $null
            $rootHwnd = $null
            $captureResult = $null

            switch ($mode) {
                "window" {
                    $rootHwnd = Resolve-WindowHandle -Payload $payload
                    $captureResult = Save-WindowCapture -Handle $rootHwnd -Path ([string]$payload.path)
                    $root = [System.Windows.Automation.AutomationElement]::FromHandle($rootHwnd)
                    break
                }
                "screen" {
                    $captureBounds = Resolve-CaptureBounds -Payload $payload
                    $captureResult = Save-ScreenCapture -Path ([string]$payload.path) -Bounds $captureBounds

                    if (($null -ne $hwndProperty -and -not [string]::IsNullOrWhiteSpace([string]$hwndProperty.Value)) -or
                        ($null -ne $titleProperty -and -not [string]::IsNullOrWhiteSpace([string]$titleProperty.Value))) {
                        $rootHwnd = Resolve-WindowHandle -Payload $payload
                        $root = [System.Windows.Automation.AutomationElement]::FromHandle($rootHwnd)
                    }
                    else {
                        $root = [System.Windows.Automation.AutomationElement]::RootElement
                    }

                    break
                }
                default {
                    throw "Unsupported see mode '$mode'"
                }
            }

            $rawElements = @(Find-AutomationElements -Root $root -Payload $payload -VisibleBounds $captureResult.bounds -DefaultMaxResults 40)
            $indexedElements = @(Add-ElementIdentifiers -Elements $rawElements)
            $normalizedElements = @()
            foreach ($item in $indexedElements) {
                $normalizedElements += ,([pscustomobject]$item)
            }

            Save-AnnotatedCapture -Path $captureResult.path -AnnotatedPath ([string]$payload.annotatedPath) -CaptureBounds $captureResult.bounds -Elements $indexedElements

            $targetHwndText = $null
            $targetTitle = $null
            if ($rootHwnd -ne $null) {
                $targetHwndText = ("0x{0:X}" -f $rootHwnd.ToInt64())
                $targetTitle = [PeekabooWindows.NativeMethods]::GetWindowTitle($rootHwnd)
            }

            Write-JsonResult -Ok $true -Result @{
                mode         = $mode
                path         = $captureResult.path
                annotatedPath = [string]$payload.annotatedPath
                bounds       = $captureResult.bounds
                target       = @{
                    hwnd  = $targetHwndText
                    title = $targetTitle
                }
                elements     = $normalizedElements
                elementCount = $normalizedElements.Count
            }
            break
        }
        "ui-find" {
            $root = Resolve-AutomationRoot -Payload $payload
            $results = @(Find-AutomationElements -Root $root -Payload $payload)
            $normalizedResults = @()
            foreach ($item in $results) {
                $normalizedResults += ,([pscustomobject]$item)
            }
            Write-JsonResult -Ok $true -Result @{
                elements = $normalizedResults
            }
            break
        }
        "ui-click" {
            $root = Resolve-AutomationRoot -Payload $payload
            $results = @(Find-AutomationElements -Root $root -Payload $payload)
            if ($results.Count -lt 1) {
                throw "No matching UI Automation element found"
            }

            $rawElements = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
            $target = $null
            $wantedName = [string]$results[0]["name"]
            $wantedAutomationId = [string]$results[0]["automationId"]
            $wantedClassName = [string]$results[0]["className"]
            $wantedControlType = [string]$results[0]["controlType"]

            foreach ($candidate in $rawElements) {
                if ([string]$candidate.Current.Name -eq $wantedName -and
                    [string]$candidate.Current.AutomationId -eq $wantedAutomationId -and
                    [string]$candidate.Current.ClassName -eq $wantedClassName -and
                    (Get-FriendlyControlTypeName -ControlType $candidate.Current.ControlType) -eq $wantedControlType) {
                    $target = $candidate
                    break
                }
            }

            if ($null -eq $target) {
                throw "Matched element could not be resolved for clicking"
            }

            $clickMethod = Invoke-AutomationElementClick -Element $target
            Write-JsonResult -Ok $true -Result @{
                clickMethod = $clickMethod
                element     = $results[0]
            }
            break
        }
        default {
            throw "Unknown action '$action'"
        }
    }
}
catch {
    Write-JsonResult -Ok $false -ErrorMessage $_.Exception.Message
    exit 1
}
