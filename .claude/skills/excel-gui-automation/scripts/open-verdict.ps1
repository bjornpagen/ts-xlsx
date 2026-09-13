<#
.SYNOPSIS
  Open a .xlsx in a VISIBLE, INTERACTIVE desktop Excel and report the verdict
  the headless COM oracle cannot see: does Excel open it clean, prompt to
  repair, auto-repair, warn on a format mismatch, or refuse it outright.

.DESCRIPTION
  tools/excel-oracle/observe.ps1 opens files with DisplayAlerts=$false, which
  SUPPRESSES the modal repair dialog - it can infer a repair happened (Open
  throws, or the name carries "[Repaired]") but never reproduces the dialog a
  real user sees. This script is the interactive counterpart: it launches Excel
  with /x (a separate process, for clean PID isolation), watches that PID's own
  top-level windows via UI Automation, classifies the first decisive state,
  screenshots it, and emits ONE JSON verdict blob on stdout.

  It is a PROBE, not a test (same contract as excel-oracle): its output is a
  recorded fact that SEEDS a corpus case. Never wire it into CI - it needs a
  licensed Excel, an interactive desktop session, and is inherently slower and
  less deterministic than the Node corpus runner.

  By default it is READ-ONLY: on a repair/mismatch prompt it clicks the safe
  negative ('No') so nothing is rewritten. Pass -AcceptRepair to click 'Yes',
  dismiss the Repairs dialog Excel then holds modal, report the repair log it
  writes, and (with -SaveRepairedTo) persist Excel's recovered output as a fixture.

.PARAMETER Path
  The .xlsx to open.

.PARAMETER AcceptRepair
  Click 'Yes' on a repair/format-mismatch prompt instead of 'No'. Reports the
  repaired workbook title and the repair log Excel writes to %TEMP%, whose
  <removedRecord> and <repairedRecord> entries say what the repair changed.
  Off by default (read-only).

.PARAMETER SaveRepairedTo
  With -AcceptRepair, save the recovered workbook to this path so you can commit
  Excel's own canonicalization of the repaired content as a fixture.

.PARAMETER TimeoutSec
  Max seconds to wait for a decisive window state before giving up. Default 40.

.PARAMETER CloseAfter
  Quit the Excel instance this run spawned when done. Off by default so you can
  inspect the window; when off, remember it leaves an EXCEL.EXE running.

.PARAMETER Shot
  Screenshot output path (defaults to a temp file - never the repo/skill dir).

.EXAMPLE
  # Read-only verdict on a suspect file
  & open-verdict.ps1 -Path .\test\corpus\fixtures\suspect.xlsx -Shot .tmp\verdict.png

.EXAMPLE
  # Accept the repair, keep Excel's recovered output, capture the log
  & open-verdict.ps1 -Path .\broken.xlsx -AcceptRepair -SaveRepairedTo .\recovered.xlsx
#>
param(
  [Parameter(Mandatory)][string]$Path,
  [switch]$AcceptRepair,
  [string]$SaveRepairedTo = '',
  [int]$TimeoutSec = 40,
  [switch]$CloseAfter,
  [string]$Shot = (Join-Path ([IO.Path]::GetTempPath()) 'xl-open-verdict.png')
)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\xl-window-lib.ps1"
. "$PSScriptRoot\uia-lib.ps1"
. "$PSScriptRoot\rot-lib.ps1"

$Path = (Resolve-Path -LiteralPath $Path).Path
$stem = [IO.Path]::GetFileNameWithoutExtension($Path)

# Classification patterns (case-insensitive substring on a window's combined
# visible text). These are Excel's English-locale strings; on first real run,
# -Dump the actual dialog text and widen these if the wording differs by build.
$reRepairPrompt = 'we found a problem with some content|recover as much as we can'
$reMismatch     = "file format and extension of.*don't match|format and extension .* don't match"
$reRejected     = 'cannot open the file|file format or file extension is not valid|is corrupt and cannot be opened|unable to read (the )?file'
# Build 20326 titles a repaired workbook's frame '<file>  -  Repaired - Excel'; older builds wrote '[Repaired]'.
$reRepairedTitle = '\[Repaired\]|\s-\s+Repaired\s+-'
# The dialog Excel holds modal after a repair: an SDM dialog titled "Repairs to '<file>'", whose list
# and buttons UI Automation does not expose, so it is found by title and closed with WM_CLOSE.
$reRepairsDialog = "^Repairs to '"
$WM_CLOSE = 0x0010

# Only ever touch the EXCEL.EXE this run spawned: the process Start-Process returns, never "every Excel
# that appeared since", which also caught the instances PowerPoint starts for its charts, so -CloseAfter
# killed them.
$others = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
if ($others -gt 0) {
  Write-Host "WARN: $others EXCEL.EXE already running; using /x to spawn an isolated instance."
}

# /x = separate process (clean PID isolation). Quote the path as one argument.
$spawned = Start-Process -FilePath 'excel.exe' -ArgumentList '/x', "`"$Path`"" -PassThru

$verdict = [ordered]@{
  path         = $Path
  verdict      = 'unknown'   # clean | repaired | repair-prompt | format-mismatch | rejected | timeout
  windowTitle  = $null
  dialogText   = $null
  accepted     = [bool]$AcceptRepair
  repairLog    = @()
  repairedPath = $null
  spawnedPid   = $null
  screenshot   = $Shot
}

function Get-SpawnedExcelPids {
  @(Get-Process -Id $spawned.Id -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
}

# Poll for the first decisive window state.
$deadline = (Get-Date).AddSeconds($TimeoutSec)
$decided = $false
$cleanStreak = 0
while ((Get-Date) -lt $deadline -and -not $decided) {
  Start-Sleep -Milliseconds 600
  $pids = Get-SpawnedExcelPids
  if (-not $pids) { continue }
  $verdict.spawnedPid = $pids[0]

  foreach ($xlPid in $pids) {
    foreach ($w in (Find-UiaWindowsByPid -ProcessId $xlPid)) {
      $title = $w.Current.Name
      $text = Get-UiaVisibleText -Scope $w

      if ($text -imatch $reRepairPrompt) {
        $verdict.verdict = 'repair-prompt'; $verdict.windowTitle = $title; $verdict.dialogText = $text; $decided = $true; break
      }
      if ($text -imatch $reMismatch) {
        $verdict.verdict = 'format-mismatch'; $verdict.windowTitle = $title; $verdict.dialogText = $text; $decided = $true; break
      }
      if ($text -imatch $reRejected) {
        $verdict.verdict = 'rejected'; $verdict.windowTitle = $title; $verdict.dialogText = $text; $decided = $true; break
      }
      # A workbook frame window (title contains the file stem) with no blocking
      # dialog. Require it to persist a couple of polls so a transient splash
      # doesn't read as 'clean'.
      if ($title -and $title -like "*$stem*") {
        if ($title -match $reRepairedTitle) {
          $verdict.verdict = 'repaired'; $verdict.windowTitle = $title; $decided = $true; break
        }
        $cleanStreak++
        if ($cleanStreak -ge 2) { $verdict.verdict = 'clean'; $verdict.windowTitle = $title; $decided = $true; break }
      }
    }
    if ($decided) { break }
  }
}
if (-not $decided) { $verdict.verdict = 'timeout' }

# Foreground + screenshot the spawned instance for the visual record.
try {
  if ($verdict.spawnedPid) {
    $hwnd = Set-ExcelForeground -TargetPid $verdict.spawnedPid
    $verdict.screenshot = $Shot
    Save-WindowScreenshot -Hwnd $hwnd -Path $Shot | Out-Null
  }
} catch { Write-Host "SCREENSHOT SKIPPED: $($_.Exception.Message)" }

# Resolve a prompt.
if ($verdict.verdict -in @('repair-prompt', 'format-mismatch')) {
  $pids = Get-SpawnedExcelPids
  $scope = $null
  foreach ($xlPid in $pids) {
    foreach ($w in (Find-UiaWindowsByPid -ProcessId $xlPid)) {
      $t = Get-UiaVisibleText -Scope $w
      if ($t -imatch $reRepairPrompt -or $t -imatch $reMismatch) { $scope = $w; break }
    }
    if ($scope) { break }
  }
  if ($scope) {
    $btn = if ($AcceptRepair) { 'Yes' } else { 'No' }
    Invoke-UiaElement -Name $btn -ControlType $script:CT::Button -Scope $scope | Out-Null

    if ($AcceptRepair) {
      # The Repairs dialog comes up a moment after 'Yes' and holds Excel modal: every COM call is
      # rejected with RPC_E_CALL_REJECTED until it closes. Only that window is closed. Clicking 'Close'
      # in every window of the instance hit the workbook frame's own close button instead.
      $noticeDeadline = (Get-Date).AddSeconds(20)
      $dismissed = $false
      while (-not $dismissed -and (Get-Date) -lt $noticeDeadline) {
        Start-Sleep -Milliseconds 600
        foreach ($xlPid in (Get-SpawnedExcelPids)) {
          foreach ($w in (Find-UiaWindowsByPid -ProcessId $xlPid)) {
            $title = $w.Current.Name
            # The verdict stays what the open did; the title records that the accepted repair ran.
            if ($title -match $reRepairedTitle) { $verdict.windowTitle = $title }
            if ($title -notmatch $reRepairsDialog) { continue }
            $verdict.dialogText = $title
            [Win32Gui]::PostMessage([IntPtr]$w.Current.NativeWindowHandle, $WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            $dismissed = $true
          }
        }
      }
      if (-not $dismissed) { Write-Host 'REPAIRS DIALOG NOT FOUND: Excel may still be modal' }

      if ($SaveRepairedTo) {
        try {
          # Through the ROT, which reaches the /x instance this run spawned rather than whichever Excel
          # registered itself first, and works under pwsh 7 (see rot-lib.ps1).
          $wb = Invoke-WhenExcelAccepts { Get-RunningWorkbook -NameLike "*$stem*" }
          if ($null -eq $wb) {
            Write-Host "SAVE-REPAIRED SKIPPED: no running workbook matches *$stem*; the ROT holds: $((Get-RunningObjectNames) -join ' | ')"
          } else {
            $SaveRepairedTo = [IO.Path]::GetFullPath($SaveRepairedTo)
            Invoke-WhenExcelAccepts { $wb.SaveAs($SaveRepairedTo, 51) }
            $verdict.repairedPath = $SaveRepairedTo
            Write-Host "SAVED REPAIRED: $SaveRepairedTo"
          }
        } catch { Write-Host "SAVE-REPAIRED SKIPPED: $($_.Exception.Message)" }
      }
      # Excel writes the repair log to %TEMP% as error<pid><n>_01.xml, named for the instance that
      # repaired, when it repairs; a later repair by a process reusing the PID overwrites it.
      if ($verdict.spawnedPid) {
        $verdict.repairLog = @(Get-ChildItem -Path $env:TEMP -Filter "error$($verdict.spawnedPid)*.xml" -ErrorAction SilentlyContinue |
          Select-Object -ExpandProperty FullName)
      }
    }
  }
}

if ($CloseAfter) {
  foreach ($xlPid in (Get-SpawnedExcelPids)) {
    Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue
  }
}

$verdict | ConvertTo-Json -Depth 5
