# Authors this case's two fixtures with real Excel: a one-variable column data table over the input cell
# D1, saved as it was built (`before.xlsx`) and saved again after row 1, the input cell's row, was deleted
# (`input-row-deleted.xlsx`). The second is Excel's own answer to what a data table looks like once its
# input cell is gone.
#
# Requires Windows + desktop Excel. Not run by CI; re-run by hand only when a fixture must change:
#   pwsh -File test/corpus/fixtures/data-table-input-cell-deleted/author.ps1

$ErrorActionPreference = 'Stop'
$outDir = $PSScriptRoot
$xl = New-Object -ComObject Excel.Application
try {
  if ($xl.Workbooks.Count -ne 0) { $xl = $null; throw 'attached to a live Excel; refusing' }
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  foreach ($variant in 'before', 'input-row-deleted') {
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Range('D1').Value2 = [double]3          # the input cell, above and to the right of the table
    $ws.Range('B3').Formula = '=D1*2'           # the formula row of a one-variable column table
    $ws.Range('A4').Value2 = [double]1
    $ws.Range('A5').Value2 = [double]2
    $ws.Range('A6').Value2 = [double]5
    [void]$ws.Range('A3:B6').Table([Type]::Missing, $ws.Range('D1'))
    if ($variant -eq 'input-row-deleted') { [void]$ws.Rows.Item(1).Delete() }
    $path = Join-Path $outDir "$variant.xlsx"
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    $wb.SaveAs($path, 51)
    $wb.Close($false)
  }
} finally {
  if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
  [GC]::Collect()
}
