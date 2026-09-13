# Dot-source this file: . '...\xl-window-lib.ps1'
# Shared foreground/screenshot/keytip primitives for driving desktop Excel's GUI.
# Ported from the verified PowerPoint equivalent (ppt-window-lib.ps1); the Win32
# foreground-lock workaround, window-rect-scoped capture, and sequential-KeyTip
# rule are app-agnostic. Only the process name (EXCEL) differs.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ("Win32Gui" -as [type])) {
  Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Win32Gui {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool f);
  [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
  public struct RECT { public int Left, Top, Right, Bottom; }
}
"@
}

function Get-ExcelHwnd {
  # $xl.Hwnd (late-bound COM property) is unreliable in this environment - read
  # the HWND from the process instead. Optionally pin to one PID when several
  # EXCEL.EXE instances are running (e.g. an open-verdict run that just spawned
  # its own), so this never foregrounds an instance you did not start.
  param([int]$TargetPid = 0)
  $procs = Get-Process EXCEL -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 }
  if ($TargetPid -ne 0) { $procs = $procs | Where-Object { $_.Id -eq $TargetPid } }
  $proc = $procs | Select-Object -First 1
  if ($proc) { return $proc.MainWindowHandle }
  return [IntPtr]::Zero
}

function Set-ExcelForeground {
  # Plain SetForegroundWindow is refused by Windows' foreground-lock timeout
  # when called from a background/detached terminal. AttachThreadInput merges
  # this thread's input state with the current foreground thread's just long
  # enough for SetForegroundWindow to be honored.
  param([IntPtr]$Hwnd = ([IntPtr]::Zero), [int]$TargetPid = 0)
  if ($Hwnd -eq [IntPtr]::Zero) { $Hwnd = Get-ExcelHwnd -TargetPid $TargetPid }
  if ($Hwnd -eq [IntPtr]::Zero) { throw "No Excel window handle found (is EXCEL.EXE running with a visible window?)" }
  $fg = [Win32Gui]::GetForegroundWindow()
  $cur = [Win32Gui]::GetCurrentThreadId()
  $pidOut = 0
  $fgThread = [Win32Gui]::GetWindowThreadProcessId($fg, [ref]$pidOut)
  [Win32Gui]::AttachThreadInput($fgThread, $cur, $true) | Out-Null
  [Win32Gui]::ShowWindow($Hwnd, 3) | Out-Null   # SW_MAXIMIZE
  [Win32Gui]::BringWindowToTop($Hwnd) | Out-Null
  [Win32Gui]::SetForegroundWindow($Hwnd) | Out-Null
  [Win32Gui]::AttachThreadInput($fgThread, $cur, $false) | Out-Null
  Start-Sleep -Milliseconds 700
  return $Hwnd
}

function Save-WindowScreenshot {
  # Captures ONLY the given window's rect, not the whole virtual screen -
  # a full-desktop capture leaks whatever else is open (other apps, chat
  # panes, unrelated windows) into the image. Falls back to the virtual
  # screen only if the rect comes back degenerate.
  param([Parameter(Mandatory)][IntPtr]$Hwnd, [Parameter(Mandatory)][string]$Path)
  $r = New-Object Win32Gui+RECT
  [Win32Gui]::GetWindowRect($Hwnd, [ref]$r) | Out-Null
  $w = $r.Right - $r.Left; $h = $r.Bottom - $r.Top
  $x = $r.Left; $y = $r.Top
  if ($w -le 0 -or $h -le 0) {
    $vs = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $x = $vs.X; $y = $vs.Y; $w = $vs.Width; $h = $vs.Height
  }
  $bmp = New-Object System.Drawing.Bitmap($w, $h)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($x, $y, 0, 0, $bmp.Size)
  $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose()
  return "${w}x${h}"
}

function Send-KeyTipSequence {
  # Ribbon KeyTips MUST be sent as SEPARATE sequential SendKeys calls with a
  # settle gap - e.g. '%' then 'N' then 'V'. Sending them as one chord ("%N")
  # fires the OLD Alt+letter accelerator instead of walking the KeyTip overlay.
  # A too-short gap (~600ms) can also miss a submenu's KeyTips before they've
  # armed; 700-900ms was reliable for Office ribbons. Prefix with '{ESC}';'{ESC}'
  # to clear any stray menu/dialog state left over from a previous attempt.
  param([Parameter(Mandatory)][string[]]$Keys, [int]$GapMs = 700)
  foreach ($k in $Keys) {
    if ($null -eq $k -or $k -eq '') { continue }
    [System.Windows.Forms.SendKeys]::SendWait($k)
    Write-Host "KEYTIP: '$k'"
    Start-Sleep -Milliseconds $GapMs
  }
}
