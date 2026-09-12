# Authors this directory's two fixtures with real Excel from one in-memory workbook: `excel-saved.xlsx`
# and `excel-saved.xlsb`. The workbook holds the errors Excel stores beside a rich value, computed and
# pasted as values, beside dynamic-array, Ctrl+Shift+Enter and plain formulas, so the pair serves both
# the rich-value error case and the binary reader's comparison with its XML twin.
#
# Requires Windows + desktop Excel. Not run by CI; re-run by hand only when the fixtures must change:
#   pwsh -File test/corpus/fixtures/an-error-value-is-stored-as-excel-stores-it/author.ps1

$ErrorActionPreference = 'Stop'
$xlsx = Join-Path $PSScriptRoot 'excel-saved.xlsx'
$xlsb = Join-Path $PSScriptRoot 'excel-saved.xlsb'
Remove-Item -Force -ErrorAction SilentlyContinue $xlsx, $xlsb

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Range('D1').Value2 = 1
  $ws.Range('D2').Value2 = 2
  $ws.Range('D3').Value2 = 3
  $ws.Range('A2').Value2 = 'blocker'
  $ws.Range('A1').Formula2 = '=SEQUENCE(3)'                 # #SPILL!, blocked by A2
  $ws.Range('C1').Formula2 = '=FILTER(D1:D3,D1:D3>100)'     # #CALC!, nothing matches
  $ws.Range('E1').Formula2 = '=FIELDVALUE(1,"x")'           # #FIELD!
  $ws.Range('I1').Formula2 = '=LAMBDA(x,x)'                 # #CALC!, a lambda never called
  $ws.Range('J2:K2').Merge()
  $ws.Range('J1').Formula2 = '=SEQUENCE(3)'                 # #SPILL!, blocked by a merge
  $ws.Range('A1048576').Formula2 = '=SEQUENCE(3)'           # #SPILL!, off the sheet's edge
  $ws.Range('A1').Copy() | Out-Null
  $ws.Range('G1').PasteSpecial(-4163) | Out-Null            # xlPasteValues
  $ws.Range('C1').Copy() | Out-Null
  $ws.Range('G2').PasteSpecial(-4163) | Out-Null
  $excel.CutCopyMode = $false
  $ws.Range('M1').Formula2 = '=SEQUENCE(3)'                 # a dynamic array spilling over M1:M3
  $ws.Range('O1').Formula2 = '=SUM(D1:D3*2)'                # a dynamic array returning one value
  $ws.Range('Q1:Q3').FormulaArray = '=D1:D3*2'              # a Ctrl+Shift+Enter formula over Q1:Q3
  $ws.Range('S1').Formula = '=D1*2'                         # a plain formula
  $wb.SaveAs($xlsx, 51)   # xlOpenXMLWorkbook
  $wb.SaveAs($xlsb, 50)   # xlExcel12
  $wb.Close($false)
} finally {
  $excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
Get-Item $xlsx, $xlsb | Select-Object Name, Length
