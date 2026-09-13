# Excel-oracle harness: geometry readback probe.
#
# Sibling of observe.ps1. Where that one asks "what formula/value does Excel see in these cells?",
# this one asks "what does Excel think the GEOMETRY of this sheet is?": per-row RowHeight, per-column
# ColumnWidth, and the sheet's StandardHeight/StandardWidth, then optionally re-saves a copy so the
# caller can read back the ht/width Excel itself considers canonical.
#
# That readback is how a clamp is distinguished from a passthrough. A package can carry an ht/width
# above what Excel's own object model accepts (`ht="5000"`); only reading the value back and diffing
# Excel's re-save shows whether it was clamped (rows: saturate at 409.6), quantized (widths: to
# 1/256 char), or honoured verbatim. See docs/knowledge/specs/grid-geometry-limits-are-excels-not-the-schemas.md.
#
# It takes an EXISTING workbook path rather than a probe spec: the fixtures worth asking this about
# are ones a writer already produced. Emit them however you like, then point this at the file.
#
# Every guardrail of observe.ps1 applies verbatim and for the same reason, since a stray modal here
# deadlocks the agent forever: DisplayAlerts=$false + AutomationSecurity=ForceDisable +
# AskToUpdateLinks=$false, the COM work inside a background job under a wall-clock watchdog
# (../excel-com-job.ps1) that kills only the EXCEL.EXE THIS run started, and a finally that always Quit()s.
#
# Automation-open is not interactive-open. Those same suppressed modals mean this probe can detect
# that a repair happened (Open threw / the name carries "[Repaired]") but never reports the
# interactive open verdict. For that, escalate to the excel-gui-automation skill. See ADR 0013.
#
# Usage:  pwsh -NoProfile -File tools/excel-oracle/read-geometry.ps1 -Path <file.xlsx> [-Rows 4] [-Cols 4]
#                [-SaveAsPath <out.xlsx>] [-NoResave] [-TimeoutSec 120]
# Emits ONE JSON observation blob on stdout.

param(
  [Parameter(Mandatory = $true)] [string] $Path,
  [int] $Rows = 4,
  [int] $Cols = 4,
  [string] $SaveAsPath = '',
  [switch] $NoResave,
  [int] $TimeoutSec = 120
)

$ErrorActionPreference = 'Stop'

# Refuse on a host without Excel rather than emitting an empty observation that reads like a finding.
# (observe.ps1 gets this guard from run.ts; this probe is invoked directly, so it carries its own.)
if (-not [Type]::GetTypeFromProgID('Excel.Application')) {
  Write-Error 'read-geometry.ps1: no registered Excel COM server (ProgID Excel.Application). This probe requires Excel Desktop on this Windows host.'
  exit 3
}

$Path = (Resolve-Path -LiteralPath $Path).Path
$resave = -not $NoResave
if ($resave) {
  if ($SaveAsPath -eq '') {
    $dir = [IO.Path]::GetDirectoryName($Path)
    $stem = [IO.Path]::GetFileNameWithoutExtension($Path)
    $SaveAsPath = Join-Path $dir "$stem.excel-resaved.xlsx"
  }
  # Absolutize before handing it to COM. Excel resolves a relative SaveAs path against ITS OWN
  # default file location, not this shell's cwd, and fails with a misleading "unable to get the
  # SaveAs property" rather than writing to the wrong place.
  $SaveAsPath = [IO.Path]::GetFullPath((Join-Path (Get-Location).Path $SaveAsPath))
}

. "$PSScriptRoot/../excel-com-job.ps1"

$work = {
  param($Path, $Rows, $Cols, $SaveAsPath, $Resave)

  $msoAutomationSecurityForceDisable = 3
  $xlOpenXMLWorkbook = 51

  $result = [ordered]@{
    version        = $null
    build          = $null
    openThrew      = $false
    openError      = $null
    repaired       = $null
    workbookName   = $null
    standardHeight = $null
    standardWidth  = $null
    rowHeights     = @()
    colWidths      = @()
    resaved        = [bool]$Resave
    resavedPath    = $(if ($Resave) { $SaveAsPath } else { $null })
    resaveThrew    = $false
    resaveError    = $null
  }

  $excel = $null
  $wb = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    Register-ExcelProcess $excel  # first, before any call that can hang: the watchdog kills only this
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable
    $excel.AskToUpdateLinks = $false
    $excel.AlertBeforeOverwriting = $false

    $result.version = $excel.Version
    try { $result.build = $excel.Build } catch { $result.build = $null }

    try {
      # ReadOnly=$true: the observed file is a fixture under measurement and must come out of the
      # run byte-identical. The canonicalized geometry is what SaveAs writes to a separate copy.
      $wb = $excel.Workbooks.Open($Path, 0, $true)
    } catch {
      $result.openThrew = $true
      $result.openError = $_.Exception.Message
    }

    if ($null -ne $wb) {
      $result.workbookName = $wb.Name
      $result.repaired = ($wb.Name -match '\[Repaired\]')

      $sheet = $wb.Worksheets.Item(1)
      $result.standardHeight = [double]$sheet.StandardHeight
      $result.standardWidth = [double]$sheet.StandardWidth

      for ($r = 1; $r -le $Rows; $r++) {
        $result.rowHeights += [ordered]@{ row = $r; height = [double]$sheet.Rows.Item($r).RowHeight }
      }
      for ($c = 1; $c -le $Cols; $c++) {
        $result.colWidths += [ordered]@{ col = $c; width = [double]$sheet.Columns.Item($c).ColumnWidth }
      }

      if ($Resave) {
        try {
          if (Test-Path -LiteralPath $SaveAsPath) { Remove-Item -LiteralPath $SaveAsPath -Force }
          $wb.SaveAs($SaveAsPath, $xlOpenXMLWorkbook)
        } catch {
          $result.resaveThrew = $true
          $result.resaveError = $_.Exception.Message
        }
      }
    }
  } finally {
    if ($null -ne $wb) {
      try { $wb.Close($false) } catch {}
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    if ($null -ne $excel) {
      try { $excel.Quit() } catch {}
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }

  $result | ConvertTo-Json -Depth 6 -Compress
}

Invoke-ExcelComJob -Work $work -ArgumentList $Path, $Rows, $Cols, $SaveAsPath, $resave -TimeoutSec $TimeoutSec
