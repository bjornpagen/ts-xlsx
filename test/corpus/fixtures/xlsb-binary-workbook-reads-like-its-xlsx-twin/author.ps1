# Authors this case's two fixtures with real Excel, so both are genuine Excel output rather than
# something this library invented: `source.xlsb` (binary BIFF12) and `source.xlsx` (OOXML XML), saved
# from one and the same in-memory workbook. That pairing is the case's whole point: the XML twin is
# an independent oracle for what the binary must decode to.
#
# Requires Windows + desktop Excel. Not run by CI; re-run by hand only when the fixture must change:
#   pwsh -File test/corpus/fixtures/xlsb-binary-workbook-reads-like-its-xlsx-twin/author.ps1

$ErrorActionPreference = 'Stop'
$outDir = $PSScriptRoot
$xlsb = Join-Path $outDir 'source.xlsb'
$xlsx = Join-Path $outDir 'source.xlsx'
Remove-Item -Force -ErrorAction SilentlyContinue $xlsb, $xlsx

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

  # --- Sheet 1: every value kind and every style facet the reader claims to decode ---
  $values = $wb.Worksheets.Item(1)
  $values.Name = 'Values'
  $values.Range('A1').Value2 = 'kind'
  $values.Range('B1').Value2 = 'value'
  $values.Range('A2').Value2 = 'small int'      # RK, integer form
  $values.Range('B2').Value2 = 10
  $values.Range('A3').Value2 = 'two decimals'   # RK, fX100 form
  $values.Range('B3').Value2 = 1.23
  $values.Range('A4').Value2 = 'wide float'     # too wide for RK -> BrtCellReal
  $values.Range('B4').Value2 = 1234.5678
  $values.Range('A5').Value2 = 'negative'
  $values.Range('B5').Value2 = -42
  $values.Range('A6').Value2 = 'boolean'
  $values.Range('B6').Value2 = $true
  $values.Range('A7').Value2 = 'error'
  $values.Range('B7').Formula = '=1/0'
  $values.Range('A8').Value2 = 'date'
  $values.Range('B8').Formula = '=DATE(2020,1,2)'
  $values.Range('B8').NumberFormat = 'yyyy-mm-dd'
  $values.Range('A9').Value2 = 'formula number'
  $values.Range('B9').Formula = '=SUM(B2:B4)'
  $values.Range('A10').Value2 = 'formula text'
  $values.Range('B10').Formula = '=UPPER("widget")'
  $values.Range('A11').Value2 = 'formula bool'
  $values.Range('B11').Formula = '=1>0'
  $values.Range('A12').Value2 = 'unicode'
  # The dash is deliberate cell content, not prose: the case asserts the binary reader
  # hands it back verbatim, so it has to survive here as the character Excel will store.
  # charcheck-disable-next-line no-em-dash-in-source
  $values.Range('B12').Value2 = 'naïve — 日本語'
  $values.Range('A13').Value2 = 'builtin numfmt'
  $values.Range('B13').Value2 = 0.125
  $values.Range('B13').NumberFormat = '0.00%'
  $values.Range('A14').Value2 = 'styled blank'   # formatted but empty -> BrtCellBlank
  $values.Range('B14').Interior.Color = 0x00FFFF

  # Header row: bold, sized, themed text on an explicit RGB fill.
  $values.Range('A1:B1').Font.Bold = $true
  $values.Range('A1:B1').Font.Size = 13
  $values.Range('A1:B1').Interior.Color = 0xEEDDCC

  # One cell carrying every non-default alignment/protection/border facet at once.
  $facets = $values.Range('D2')
  $facets.Value2 = 'facets'
  $facets.Font.Italic = $true
  $facets.Font.Name = 'Consolas'
  $facets.Font.Color = 0x0000C0            # BGR -> ARGB FFC00000
  $facets.Font.Underline = 2               # xlUnderlineStyleSingle
  $facets.HorizontalAlignment = -4108      # xlCenter
  $facets.VerticalAlignment = -4160        # xlTop
  $facets.WrapText = $true
  $facets.IndentLevel = 2
  $facets.Locked = $false
  $facets.Borders.Item(7).LineStyle = 1    # xlEdgeLeft, xlContinuous
  $facets.Borders.Item(7).Weight = 2       # xlThin
  $facets.Borders.Item(9).LineStyle = -4115  # xlEdgeBottom, xlDash
  $values.Range('D4').Value2 = 'rotated'
  $values.Range('D4').Orientation = 45
  $values.Range('D6').Value2 = 'hatched'
  $values.Range('D6').Interior.Pattern = 16 # xlPatternGray16 -> a non-solid pattern fill
  $values.Range('F2:G3').Merge()
  $values.Range('F2').Value2 = 'merged'

  # --- Sheet 2: the grid layout facets (column/row geometry and inherited formats) ---
  $grid = $wb.Worksheets.Add()
  $grid.Name = 'Grid'
  $grid.Range('A1').Value2 = 'a'
  $grid.Range('B1').Value2 = 'b'
  $grid.Range('C1').Value2 = 'c'
  $grid.Range('E1').Value2 = 'e'
  $grid.Columns.Item(1).ColumnWidth = 24
  $grid.Columns.Item(2).Hidden = $true
  $grid.Columns.Item(3).Font.Bold = $true      # a column-level default format
  $grid.Range('C:D').Group() | Out-Null        # an outlined column span
  $grid.Rows.Item(1).RowHeight = 30
  $grid.Range('A3').Value2 = 'hidden row'
  $grid.Rows.Item(3).Hidden = $true
  $grid.Range('A5').Value2 = 'outlined'
  $grid.Range('5:6').Group() | Out-Null
  $grid.Rows.Item(8).Font.Italic = $true       # a row-level default format
  $grid.Range('A8').Value2 = 'row format'
  # Protected with a password, so the binary sheet carries the agile credential record beside the flags,
  # and with formatting cells, sorting and filtering allowed and objects left editable.
  # Protect(Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly, AllowFormattingCells,
  #   AllowFormattingColumns, AllowFormattingRows, AllowInsertingColumns, AllowInsertingRows,
  #   AllowInsertingHyperlinks, AllowDeletingColumns, AllowDeletingRows, AllowSorting, AllowFiltering)
  $grid.Protect('abc', $false, $true, $true, $false, $true, $false, $false, $false, $false, $false, $false, $false, $true, $true)

  # --- Sheet 3: hidden, to prove sheet order/visibility survive ---
  $quiet = $wb.Worksheets.Add()
  $quiet.Name = 'Quiet'
  $quiet.Range('A1').Value2 = 'shh'
  $quiet.Visible = 0   # xlSheetHidden
  # Protected with no password and no cell selectable, the two flags a protected sheet permits by default.
  $quiet.Protect()
  $quiet.EnableSelection = -4142   # xlNoSelection

  $wb.Worksheets.Item('Values').Activate()
  $wb.SaveAs($xlsb, 50)   # xlExcel12
  $wb.SaveAs($xlsx, 51)   # xlOpenXMLWorkbook
  $wb.Close($false)
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Get-Item $xlsb, $xlsx | Select-Object Name, Length
