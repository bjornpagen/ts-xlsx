# Authors `excel-saved.xlsx` with real Excel: a report sheet whose one drawing holds a clustered column
# chart over A1:B4 and a rectangle shape, the kind of content the library keeps byte for byte rather
# than models, so a picture added to the sheet has to join that drawing.
#
# Requires Windows + desktop Excel. Not run by CI; re-run by hand only when the fixture must change:
#   pwsh -File test/corpus/fixtures/picture-joins-a-kept-drawing/author.ps1

$ErrorActionPreference = 'Stop'
$xlsx = Join-Path $PSScriptRoot 'excel-saved.xlsx'
Remove-Item -Force -ErrorAction SilentlyContinue $xlsx

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $s = $wb.Worksheets.Item(1)
  $s.Name = 'Report'
  $s.Range('A1').Value2 = 'Quarter'
  $s.Range('B1').Value2 = 'Sales'
  foreach ($r in 2..4) {
    $s.Cells.Item($r, 1).Value2 = "Q$($r - 1)"
    $s.Cells.Item($r, 2).Value2 = $r * 10
  }

  $chart = $s.Shapes.AddChart2(201, 51, 200, 20, 300, 180)                        # xlColumnClustered
  $chart.Chart.SetSourceData($s.Range('A1:B4'))
  $shape = $s.Shapes.AddShape(1, 20, 120, 120, 40)                                # msoShapeRectangle
  $shape.TextFrame2.TextRange.Text = 'Draft'

  $wb.SaveAs($xlsx, 51)                                                           # xlOpenXMLWorkbook
  $wb.Close($false)
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Get-Item $xlsx | Select-Object Name, Length
