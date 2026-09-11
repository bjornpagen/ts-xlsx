# Authors this case's fixture with real Excel: `excel-saved.xlsx`, a sheet whose conditional formats, data
# validations and table carry formulas calling post-2007 functions, a LET, and (in the table, the one
# place Excel accepts it) a function passed as a value. What Excel writes into each `<formula>`,
# `<formula1>`, `<xm:f>`, `<cfvo val>`, `<calculatedColumnFormula>` and `<totalsRowFormula>` is the
# spelling the library must read back to plain formula text and write back unchanged.
#
# Excel refuses a function passed as a value in a conditional format or a validation formula, and a
# dynamic-array function as a list validation's source, so neither can appear in a file it saved.
#
# Requires Windows + desktop Excel. Not run by CI; re-run by hand only when the fixture must change:
#   pwsh -File test/corpus/fixtures/rule-and-table-formulas-keep-function-prefixes/author.ps1

$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot 'excel-saved.xlsx'
Remove-Item -Force -ErrorAction SilentlyContinue $out

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 2) {
    [void]$wb.Worksheets.Add([Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
  }
  while ($wb.Worksheets.Count -gt 2) { $wb.Worksheets.Item(3).Delete() }
  $s = $wb.Worksheets.Item(1)
  $s.Name = 'S1'
  $other = $wb.Worksheets.Item(2)
  $other.Name = 'S2'
  foreach ($sheet in @($s, $other)) {
    1..3 | ForEach-Object { $sheet.Cells.Item($_, 1).Value2 = $_; $sheet.Cells.Item($_, 2).Value2 = $_ }
  }

  # Conditional formats: an expression calling XLOOKUP, one binding a LET, a cell-value bound calling
  # MAXIFS, an expression reaching another sheet (which Excel stores in the x14 extension), and a
  # colour scale and a data bar whose anchors are formulas.
  $fc = $s.Range('A1:A3').FormatConditions.Add(2, [Reflection.Missing]::Value, '=XLOOKUP(A1,$B$1:$B$3,$B$1:$B$3)=1')
  $fc.Interior.Color = 255
  $fc = $s.Range('B1:B3').FormatConditions.Add(2, [Reflection.Missing]::Value, '=LET(x,B1,x>1)')
  $fc.Interior.Color = 255
  $fc = $s.Range('C1:C3').FormatConditions.Add(1, 5, '=MAXIFS($A$1:$A$3,$A$1:$A$3,">0")')
  $fc.Interior.Color = 255
  $fc = $s.Range('C4:C6').FormatConditions.Add(2, [Reflection.Missing]::Value, '=XLOOKUP(1,S2!$A$1:$A$3,S2!$B$1:$B$3)=1')
  $fc.Interior.Color = 255
  $scale = $s.Range('E1:E3').FormatConditions.AddColorScale(2)
  $scale.ColorScaleCriteria.Item(1).Type = 4
  $scale.ColorScaleCriteria.Item(1).Value = '=MINIFS($A$1:$A$3,$A$1:$A$3,">0")'
  $scale.ColorScaleCriteria.Item(2).Type = 4
  $scale.ColorScaleCriteria.Item(2).Value = '=MAXIFS($A$1:$A$3,$A$1:$A$3,">0")'
  $bar = $s.Range('F1:F3').FormatConditions.AddDatabar()
  $bar.MinPoint.Modify(4, '=MINIFS($A$1:$A$3,$A$1:$A$3,">0")')
  $bar.MaxPoint.Modify(4, '=MAXIFS($A$1:$A$3,$A$1:$A$3,">0")')

  # Data validations: a custom formula calling XLOOKUP, a decimal range between MINIFS and MAXIFS, and a
  # custom formula reaching another sheet.
  $s.Range('D1').Validation.Add(7, 1, 1, '=ISNUMBER(XLOOKUP(D1,$A$1:$A$3,$A$1:$A$3))')
  $s.Range('D2').Validation.Add(2, 1, 1, '=MINIFS($A$1:$A$3,$A$1:$A$3,">0")', '=MAXIFS($A$1:$A$3,$A$1:$A$3,">0")')
  $s.Range('D3').Validation.Add(7, 1, 1, '=ISNUMBER(XLOOKUP(D3,S2!$A$1:$A$3,S2!$A$1:$A$3))')

  # A table whose second column is calculated from XLOOKUP and LET, with a totals row whose two totals
  # are custom formulas, one passing SUM as a value.
  $s.Range('G1').Value2 = 'a'
  $s.Range('H1').Value2 = 'b'
  2..4 | ForEach-Object { $s.Cells.Item($_, 7).Value2 = $_ - 1 }
  $table = $s.ListObjects.Add(1, $s.Range('G1:H4'), [Reflection.Missing]::Value, 1)
  $table.Name = 'T'
  $s.Range('H2').Formula2 = '=XLOOKUP([@a],$A$1:$A$3,$B$1:$B$3)+LET(x,1,x)'
  $table.ShowTotals = $true
  $s.Range('G5').Formula2 = '=SUM(BYROW(T[a],SUM))'
  $s.Range('H5').Formula2 = '=SUBTOTAL(109,[b])+XLOOKUP(1,$A$1:$A$3,$B$1:$B$3)'

  $s.Activate()
  $wb.SaveAs($out, 51)
  $wb.Close($false)
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Get-Item $out | Select-Object Name, Length
