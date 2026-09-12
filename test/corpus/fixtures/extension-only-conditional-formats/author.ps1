# Authors `excel-saved.xlsx` with real Excel: one sheet of conditional formats Excel stores only in the
# 2009 x14 extension (`<x14:conditionalFormatting>` inside the worksheet `<extLst>`) beside a classic
# rule and a classic data bar that links to its extension, so a reader has to tell the two apart.
#
# - A1:A6, an expression rule reaching another sheet, with a bold font and a red fill.
# - C1:C6, a cellIs rule whose operand is on another sheet, with a blue font and a number format.
# - D1:D6, a classic same-sheet expression rule with a yellow fill: the control.
# - E1:E6, an icon set from a 2009 family (3 stars).
# - F1:F6, a custom icon set: 3 arrows whose middle icon is "no icon" and whose top icon is a flag.
# - G1:G6, a classic icon set in reverse order with its values hidden and a `>` threshold on a number.
# - H1:H6, a colour scale whose low anchor is a formula on another sheet.
# - B1:B6, a data bar with the 2010 facets: a border, right-to-left, a solid fill and no axis.
#
# COM refuses `[System.Reflection.Missing]::Value` once it is held in a variable, so it is spelled
# inline at every optional argument.
#
# Requires Windows + desktop Excel. Not run by CI; re-run by hand only when the fixture must change:
#   pwsh -File test/corpus/fixtures/extension-only-conditional-formats/author.ps1

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
  $s.Name = 'Rules'
  $data = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $s)
  $data.Name = 'Data'
  $data.Range('A1').Value2 = 3
  foreach ($r in 1..6) {
    foreach ($c in 1..8) { $s.Cells.Item($r, $c).Value2 = $r }
    $s.Cells.Item($r, 2).Value2 = $r - 3
  }

  $fc = $s.Range('A1:A6').FormatConditions.Add(2, [System.Reflection.Missing]::Value, '=A1>Data!$A$1')   # xlExpression
  $fc.Interior.Color = 255
  $fc.Font.Bold = $true

  $fc2 = $s.Range('C1:C6').FormatConditions.Add(1, 5, '=Data!$A$1')             # xlCellValue, xlGreater
  $fc2.Font.Color = 16711680
  $fc2.NumberFormat = '0.00'

  $fc3 = $s.Range('D1:D6').FormatConditions.Add(2, [System.Reflection.Missing]::Value, '=D1>4')
  $fc3.Interior.Color = 65535

  $stars = $s.Range('E1:E6').FormatConditions.AddIconSetCondition()
  $stars.IconSet = $wb.IconSets.Item(18)                                          # xl3Stars

  $custom = $s.Range('F1:F6').FormatConditions.AddIconSetCondition()
  $custom.IconSet = $wb.IconSets.Item(1)                                          # xl3Arrows
  $custom.IconCriteria.Item(2).Icon = -1                                          # xlIconNoCellIcon
  $custom.IconCriteria.Item(3).Icon = 22                                          # a flag

  $reversed = $s.Range('G1:G6').FormatConditions.AddIconSetCondition()
  $reversed.IconSet = $wb.IconSets.Item(1)
  $reversed.ReverseOrder = $true
  $reversed.ShowIconOnly = $true
  $reversed.IconCriteria.Item(2).Type = 0                                         # xlConditionValueNumber
  $reversed.IconCriteria.Item(2).Value = 2
  $reversed.IconCriteria.Item(2).Operator = 5                                     # xlGreater

  $scale = $s.Range('H1:H6').FormatConditions.AddColorScale(2)
  $scale.ColorScaleCriteria.Item(1).Type = 4                                      # xlConditionValueFormula
  $scale.ColorScaleCriteria.Item(1).Value = '=Data!$A$1'

  $bar = $s.Range('B1:B6').FormatConditions.AddDatabar()
  $bar.BarFillType = 0                                                            # solid
  $bar.BarBorder.Type = 1                                                         # solid border
  $bar.BarBorder.Color.Color = 0
  $bar.Direction = -5004                                                          # right to left
  $bar.AxisPosition = 2                                                           # no axis
  $bar.NegativeBarFormat.ColorType = 1
  $bar.ShowValue = $false

  $s.Activate()
  $wb.SaveAs($xlsx, 51)                                                           # xlOpenXMLWorkbook
  $wb.Close($false)
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Get-Item $xlsx | Select-Object Name, Length
