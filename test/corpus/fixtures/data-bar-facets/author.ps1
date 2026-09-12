# Authors `excel-saved.xlsx` with real Excel: one sheet of data bars, each exercising facets the classic
# `<dataBar>` element cannot carry, which Excel stores in the linked `<x14:dataBar>` of the worksheet
# extension.
#
# - A1:A6, automatic low and high anchors, the anchors Excel's own ribbon gives a new bar.
# - B1:B6, a solid black border, right-to-left, a solid fill, no axis and the values hidden.
# - C1:C6, negative values in their own red fill and blue border, the axis at the midpoint in green.
# - D1:D6, a green bar whose shortest and longest lengths are 20% and 80% of the cell.
# - E1:E6, a number for the low anchor and a percentile for the high one.
#
# COM refuses `[System.Reflection.Missing]::Value` once it is held in a variable; nothing here needs it.
#
# Requires Windows + desktop Excel. Not run by CI; re-run by hand only when the fixture must change:
#   pwsh -File test/corpus/fixtures/data-bar-facets/author.ps1

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
  $s.Name = 'Bars'
  foreach ($r in 1..6) {
    foreach ($c in 1..5) { $s.Cells.Item($r, $c).Value2 = $r - 3 }
  }

  $auto = $s.Range('A1:A6').FormatConditions.AddDatabar()
  $auto.MinPoint.Modify(6)                                                        # xlConditionValueAutomaticMin
  $auto.MaxPoint.Modify(7)                                                        # xlConditionValueAutomaticMax

  $framed = $s.Range('B1:B6').FormatConditions.AddDatabar()
  $framed.BarFillType = 0                                                         # xlDataBarFillSolid
  $framed.BarBorder.Type = 1                                                      # xlDataBarBorderSolid
  $framed.BarBorder.Color.Color = 0
  $framed.Direction = -5004                                                       # xlRTL
  $framed.AxisPosition = 2                                                        # xlDataBarAxisNone
  $framed.ShowValue = $false

  $negative = $s.Range('C1:C6').FormatConditions.AddDatabar()
  $negative.BarBorder.Type = 1
  $negative.NegativeBarFormat.ColorType = 0                                       # xlDataBarColor
  $negative.NegativeBarFormat.Color.Color = 255                                   # red
  $negative.NegativeBarFormat.BorderColorType = 0
  $negative.NegativeBarFormat.BorderColor.Color = 16711680                        # blue
  $negative.AxisPosition = 1                                                      # xlDataBarAxisMidpoint
  $negative.AxisColor.Color = 65280                                               # green

  $lengths = $s.Range('D1:D6').FormatConditions.AddDatabar()
  $lengths.BarColor.Color = 5296274
  $lengths.PercentMin = 20
  $lengths.PercentMax = 80

  $anchored = $s.Range('E1:E6').FormatConditions.AddDatabar()
  $anchored.MinPoint.Modify(0, -2)                                                # xlConditionValueNumber
  $anchored.MaxPoint.Modify(5, 90)                                                # xlConditionValuePercentile

  $s.Activate()
  $wb.SaveAs($xlsx, 51)                                                           # xlOpenXMLWorkbook
  $wb.Close($false)
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Get-Item $xlsx | Select-Object Name, Length
