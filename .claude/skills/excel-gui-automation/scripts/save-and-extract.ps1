<#
.SYNOPSIS
  Save the workbook you've been GUI-driving (via its already-running COM
  instance) and extract its .xlsx package to a fresh directory, so you can read
  the actual generated OOXML rather than trusting the screenshot alone. A
  screenshot proves the UI reacted; it does not prove the XML is correct or
  wired up (relationships, ids, namespaces) - always close the loop by reading
  the parts this script extracts.

.PARAMETER NameLike
  Wildcard matched against each open workbook's full path and its file name, in
  whichever running Excel holds it (useful when more than one is open).

.PARAMETER DestDir
  Extraction target. If it already exists, a fresh "<DestDir>-<random>" is used
  instead of deleting anything - do not delete-and-recreate: this sandbox's
  Bash/PowerShell guard can false-positive-block Remove-Item when the
  surrounding command text also contains regex-like substrings (e.g. a literal
  "r:" from an XML namespace or "\w+"), so a fresh directory per run is more
  reliable than reuse.
#>
param(
  [Parameter(Mandatory)][string]$NameLike,
  [Parameter(Mandatory)][string]$DestDir
)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\rot-lib.ps1"

$wb = Get-RunningWorkbook -NameLike $NameLike
if ($null -eq $wb) {
  throw "No open workbook matching '$NameLike'; the ROT holds: $((Get-RunningObjectNames) -join ' | ')"
}

$xlsxPath = $wb.FullName
$wb.Save()
Write-Host "SAVED: $xlsxPath (Sheets=$($wb.Worksheets.Count))"

$zip = "$DestDir.zip"
Copy-Item -Path $xlsxPath -Destination $zip -Force
$dir = $DestDir
if (Test-Path $dir) { $dir = "$DestDir-$(Get-Random)" }
Expand-Archive -Path $zip -DestinationPath $dir -Force
Write-Host "EXTRACT DIR: $dir"

Get-ChildItem -Path (Join-Path $dir "xl\worksheets") -Filter *.xml -ErrorAction SilentlyContinue |
  ForEach-Object { Write-Host "SHEET: $($_.Name) ($($_.Length) bytes)" }
Get-ChildItem -Path (Join-Path $dir "xl\worksheets\_rels") -Filter *.rels -ErrorAction SilentlyContinue |
  ForEach-Object { Write-Host "RELS: $($_.Name)" }
