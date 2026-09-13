# vba-compiler harness: VBIDE COM driver.
#
# Turns VBA module SOURCE into a genuinely-compiled `vbaProject.bin` (or a whole `.xlsm`) by driving a
# real, headless Excel through the VBIDE object model. This is the ONLY sound way to author/edit VBA:
# Excel does not recompile from source on open: a module ships its p-code (PerformanceCache) and Excel
# runs THAT; a from-scratch or byte-spliced project with no/mismatched p-code either throws "Invalid
# data format" or silently runs stale code (recorded finding 2026-07-24; see ../../docs/decisions).
#
# Contract mirrors tools/excel-oracle/observe.ps1: it owns EVERY guardrail, because a stray modal here
# deadlocks the caller forever:
#   - Visible=$false + DisplayAlerts=$false + AutomationSecurity=Low (macros MUST load so they compile);
#   - the COM work runs inside a background job wrapped by a wall-clock watchdog (../excel-com-job.ps1);
#   - on timeout the job is stopped and the EXCEL.EXE THIS run started, and no other, is force-killed;
#   - the job's own finally always Quit()s, ReleaseComObject()s, and GCs, whether or not it threw.
#
# PROBE/build tool, NEVER in CI: needs a licensed Excel and Trust access to the VBA project object model
# (HKCU\Software\Microsoft\Office\<ver>\Excel\Security\AccessVBOM = 1). Its output is a recorded
# artifact that seeds a committed corpus fixture (ADR 0012/0013 seed+lock split).

param(
  [Parameter(Mandatory = $true)] [string] $Spec,   # JSON: { modules:[{name,kind,source}], base?:path }
  [Parameter(Mandatory = $true)] [string] $Out,    # output path; *.bin extracts vbaProject.bin, else .xlsm
  [int] $TimeoutSec = 120
)

$ErrorActionPreference = 'Stop'
$Spec = (Resolve-Path -LiteralPath $Spec).Path
$specObj = Get-Content -LiteralPath $Spec -Raw | ConvertFrom-Json
$base = if ($specObj.PSObject.Properties.Name -contains 'base' -and $specObj.base) {
  (Resolve-Path -LiteralPath $specObj.base).Path
} else { '' }

. "$PSScriptRoot/../excel-com-job.ps1"

$work = {
  param($Modules, $Base, $Out)

  $msoAutomationSecurityLow = 1
  $xlOpenXMLWorkbookMacroEnabled = 52
  # vbext_ComponentType
  $ctStdModule = 1; $ctClassModule = 2; $ctMSForm = 3; $ctDocument = 100

  $result = [ordered]@{
    ok        = $false
    mode      = $(if ($Base) { 'in-place' } else { 'from-scratch' })
    out       = $Out
    modules   = @()
    error     = $null
  }

  $excel = $null
  $wb = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    Register-ExcelProcess $excel  # first, before any call that can hang: the watchdog kills only this
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = $msoAutomationSecurityLow

    if ($Base) {
      $wb = $excel.Workbooks.Open($Base)
    } else {
      $wb = $excel.Workbooks.Add()
    }
    $proj = $wb.VBProject

    foreach ($m in $Modules) {
      $name = [string]$m.name
      $kind = [string]$m.kind
      # A leading `Attribute VB_Name = "..."` line conflicts with the .Name we set and silently breaks
      # compilation (verified); strip only that one attribute, keep any others (document modules need
      # VB_Base/VB_Exposed/etc.).
      $src = [string]$m.source
      $src = [regex]::Replace($src, '^\s*Attribute\s+VB_Name\s*=\s*"[^"]*"\r?\n', '')

      $existing = $null
      try { $existing = $proj.VBComponents.Item($name) } catch { $existing = $null }

      if ($existing) {
        # Replace the code of an existing module (document code-behind, or a re-edit).
        $cm = $existing.CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString($src)
        $result.modules += [ordered]@{ name = $name; action = 'replaced'; kind = $kind }
      } else {
        $ct = switch ($kind) {
          'procedural' { $ctStdModule }
          'class'      { $ctClassModule }
          'designer'   { $ctMSForm }
          'document'   { throw "module '$name': a 'document' module cannot be added; it must already exist in -base" }
          default      { throw "module '$name': unknown kind '$kind'" }
        }
        $comp = $proj.VBComponents.Add($ct)
        $comp.Name = $name
        $comp.CodeModule.AddFromString($src)
        $result.modules += [ordered]@{ name = $name; action = 'added'; kind = $kind }
      }
    }

    if ($Out.ToLower().EndsWith('.bin')) {
      # Save to a temp .xlsm, then extract xl/vbaProject.bin.
      $tmpXlsm = [IO.Path]::Combine([IO.Path]::GetTempPath(), "vbacompile-$([guid]::NewGuid()).xlsm")
      $wb.SaveAs($tmpXlsm, $xlOpenXMLWorkbookMacroEnabled)
      $wb.Close($false); $wb = $null
      Add-Type -AssemblyName System.IO.Compression.FileSystem
      $zip = [IO.Compression.ZipFile]::OpenRead($tmpXlsm)
      try {
        $entry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/vbaProject.bin' } | Select-Object -First 1
        if (-not $entry) { throw 'the saved workbook has no xl/vbaProject.bin' }
        if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }
        [IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $Out, $true)
      } finally { $zip.Dispose() }
      Remove-Item -LiteralPath $tmpXlsm -Force -ErrorAction SilentlyContinue
    } else {
      if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out -Force }
      $wb.SaveAs($Out, $xlOpenXMLWorkbookMacroEnabled)
      $wb.Close($false); $wb = $null
    }
    $result.ok = $true
  } catch {
    $result.error = $_.Exception.Message
  } finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }

  $result | ConvertTo-Json -Depth 6 -Compress
}

Invoke-ExcelComJob -Work $work -ArgumentList $specObj.modules, $base, $Out -TimeoutSec $TimeoutSec
