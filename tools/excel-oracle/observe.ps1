# Excel-oracle harness: COM driver.
#
# Opens a .xlsx in headless Excel Desktop, reads formula/value per requested cell, optionally re-saves
# a copy (so the caller can read back the ref Excel itself canonicalizes), and emits ONE JSON
# observation blob on stdout. It owns EVERY safety guardrail, because a stray modal here deadlocks the
# agent forever:
#   - DisplayAlerts=$false + AutomationSecurity=ForceDisable + AskToUpdateLinks=$false suppress modals;
#   - the COM work runs inside a background job wrapped by a wall-clock watchdog (../excel-com-job.ps1);
#   - on timeout the job is stopped and the EXCEL.EXE THIS run started, and no other, is force-killed;
#   - the job's own finally always Quit()s, ReleaseComObject()s, and GCs, whether or not Open threw.
#
# Automation-open is not interactive-open: these guards suppress the modal *repair dialog*, so this
# harness can DETECT that a repair happened (Open throws / the workbook name carries "[Repaired]") but
# cannot reproduce the interactive dialog experience. Callers must record which class an observation is.

param(
  [Parameter(Mandatory = $true)] [string] $Path,
  # Comma-joined cell addresses (e.g. "B1,B2,D5"). Taken as one string and split here on purpose: a
  # [string[]] param bound via -File from an external spawn collapses to a single element, and Excel's
  # Range() reads a comma as a union operator, so a joined token silently reads one merged area, not
  # each cell. Splitting here keeps one address per readback. An address may name its sheet
  # (`Sheet2!B2`), which is also why a sheet name holding a comma cannot be observed.
  [string] $Cells = '',
  [string] $SaveAsPath = '',
  [switch] $NoResave,
  [int] $TimeoutSec = 90
)

$ErrorActionPreference = 'Stop'
$Path = (Resolve-Path -LiteralPath $Path).Path
$cellList = @($Cells -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
$resave = -not $NoResave
if ($resave -and $SaveAsPath -eq '') {
  $dir = [IO.Path]::GetDirectoryName($Path)
  $stem = [IO.Path]::GetFileNameWithoutExtension($Path)
  $SaveAsPath = Join-Path $dir "$stem.excel-resaved.xlsx"
}

. "$PSScriptRoot/../excel-com-job.ps1"

$work = {
  param($Path, $Cells, $SaveAsPath, $Resave)

  $msoAutomationSecurityForceDisable = 3
  $xlOpenXMLWorkbook = 51

  $result = [ordered]@{
    version      = $null
    build        = $null
    openThrew    = $false
    openError    = $null
    repaired     = $null
    workbookName = $null
    cells        = @()
    resaved      = [bool]$Resave
    resavedPath  = $(if ($Resave) { $SaveAsPath } else { $null })
    resaveThrew  = $false
    resaveError  = $null
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
      $wb = $excel.Workbooks.Open($Path, 0, $true)  # UpdateLinks=0, ReadOnly=$true
    } catch {
      $result.openThrew = $true
      $result.openError = $_.Exception.Message
    }

    if ($null -ne $wb) {
      $result.workbookName = $wb.Name
      $result.repaired = ($wb.Name -match '\[Repaired\]')

      foreach ($addr in $Cells) {
        # `Sheet!A1` names another sheet; a bare address is the first sheet's. The sheet name is taken
        # up to the last `!`, and a quoted one (`'My Sheet'!A1`) loses its quotes.
        $bang = $addr.LastIndexOf('!')
        if ($bang -ge 0) {
          $sheetName = $addr.Substring(0, $bang) -replace "^'(.*)'$", '$1' -replace "''", "'"
          $sheet = $wb.Worksheets.Item($sheetName)
          $local = $addr.Substring($bang + 1)
        } else {
          $sheet = $wb.Worksheets.Item(1)
          $local = $addr
        }
        $rng = $sheet.Range($local)
        $raw = $rng.Value2
        $result.cells += [ordered]@{
          address    = $addr
          hasFormula = [bool]$rng.HasFormula
          formula    = [string]$rng.Formula
          value      = "$raw"
          # What `value` cannot say. An empty string and a blank cell both stringify to "", and an error
          # is a bare Int32 (CVErr code) that looks like a number: the type and ISBLANK tell them apart,
          # and `Text` is what the cell displays.
          valueType  = $(if ($null -eq $raw) { 'null' } else { $raw.GetType().Name })
          text       = [string]$rng.Text
          isBlank    = [bool]$sheet.Evaluate("ISBLANK($local)")
        }
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

Invoke-ExcelComJob -Work $work -ArgumentList $Path, $cellList, $SaveAsPath, $resave -TimeoutSec $TimeoutSec
