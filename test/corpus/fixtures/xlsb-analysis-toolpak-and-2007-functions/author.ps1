# Authors this case's two fixtures with real Excel: `source.xlsb` and `source.xlsx`, saved from one and
# the same in-memory workbook. It holds a formula for each of the 105 functions an `.xlsb` cites by an
# index past `RTD`, where the [MS-XLS] function table ends: the Analysis ToolPak and CUBE functions
# Excel 2007 took into its engine, `IFERROR`, and the `*IFS` family. Each takes the fewest arguments
# Excel accepts, so the call token carries the arity the table must know, and a few repeat with an
# optional argument to show a variadic call keeps its count. The XML twin is the oracle for the text.
#
# Calculation is manual so a volatile function such as RANDBETWEEN caches one result for both saves,
# and no CUBE function goes looking for a connection.
#
# Requires Windows + desktop Excel. Not run by CI; re-run by hand only when the fixture must change:
#   pwsh -File test/corpus/fixtures/xlsb-analysis-toolpak-and-2007-functions/author.ps1

$ErrorActionPreference = 'Stop'
$outDir = $PSScriptRoot
$xlsb = Join-Path $outDir 'source.xlsb'
$xlsx = Join-Path $outDir 'source.xlsx'
Remove-Item -Force -ErrorAction SilentlyContinue $xlsb, $xlsx

$formulas = @(
  '=CUBEVALUE(1)', '=CUBEMEMBER(1,1)', '=CUBEMEMBERPROPERTY(1,1,1)', '=CUBERANKEDMEMBER(1,1,1)',
  '=HEX2BIN(1)', '=HEX2DEC(1)', '=HEX2OCT(1)', '=DEC2BIN(1)', '=DEC2HEX(1)', '=DEC2OCT(1)',
  '=OCT2BIN(1)', '=OCT2HEX(1)', '=OCT2DEC(1)', '=BIN2DEC(1)', '=BIN2OCT(1)', '=BIN2HEX(1)',
  '=IMSUB(1,1)', '=IMDIV(1,1)', '=IMPOWER(1,1)', '=IMABS(1)', '=IMSQRT(1)', '=IMLN(1)', '=IMLOG2(1)',
  '=IMLOG10(1)', '=IMSIN(1)', '=IMCOS(1)', '=IMEXP(1)', '=IMARGUMENT(1)', '=IMCONJUGATE(1)',
  '=IMAGINARY(1)', '=IMREAL(1)', '=COMPLEX(1,1)', '=IMSUM(1)', '=IMPRODUCT(1)', '=SERIESSUM(1,1,1,1)',
  '=FACTDOUBLE(1)', '=SQRTPI(1)', '=QUOTIENT(1,1)', '=DELTA(1)', '=GESTEP(1)', '=ISEVEN(1)', '=ISODD(1)',
  '=MROUND(1,1)', '=ERF(1)', '=ERFC(1)', '=BESSELJ(1,1)', '=BESSELK(1,1)', '=BESSELY(1,1)',
  '=BESSELI(1,1)', '=XIRR(1,1)', '=XNPV(1,1,1)', '=PRICEMAT(1,1,1,1,1)', '=YIELDMAT(1,1,1,1,1)',
  '=INTRATE(1,1,1,1)', '=RECEIVED(1,1,1,1)', '=DISC(1,1,1,1)', '=PRICEDISC(1,1,1,1)',
  '=YIELDDISC(1,1,1,1)', '=TBILLEQ(1,1,1)', '=TBILLPRICE(1,1,1)', '=TBILLYIELD(1,1,1)',
  '=PRICE(1,1,1,1,1,1)', '=YIELD(1,1,1,1,1,1)', '=DOLLARDE(1,1)', '=DOLLARFR(1,1)', '=NOMINAL(1,1)',
  '=EFFECT(1,1)', '=CUMPRINC(1,1,1,1,1,1)', '=CUMIPMT(1,1,1,1,1,1)', '=EDATE(1,1)', '=EOMONTH(1,1)',
  '=YEARFRAC(1,1)', '=COUPDAYBS(1,1,1)', '=COUPDAYS(1,1,1)', '=COUPDAYSNC(1,1,1)', '=COUPNCD(1,1,1)',
  '=COUPNUM(1,1,1)', '=COUPPCD(1,1,1)', '=DURATION(1,1,1,1,1)', '=MDURATION(1,1,1,1,1)',
  '=ODDLPRICE(1,1,1,1,1,1,1)', '=ODDLYIELD(1,1,1,1,1,1,1)', '=ODDFPRICE(1,1,1,1,1,1,1,1)',
  '=ODDFYIELD(1,1,1,1,1,1,1,1)', '=RANDBETWEEN(1,1)', '=WEEKNUM(1)', '=AMORDEGRC(1,1,1,1,1,1)',
  '=AMORLINC(1,1,1,1,1,1)', '=CONVERT(1,1,1)', '=ACCRINT(1,1,1,1,1,1)', '=ACCRINTM(1,1,1,1)',
  '=WORKDAY(1,1)', '=NETWORKDAYS(1,1)', '=GCD(1)', '=MULTINOMIAL(1)', '=LCM(1)', '=FVSCHEDULE(1,1)',
  '=CUBEKPIMEMBER(1,1,1)', '=CUBESET(1,1)', '=CUBESETCOUNT(1)', '=IFERROR(A1/0,0)',
  '=COUNTIFS(A1:A5,1)', '=SUMIFS(A1:A5,A1:A5,1)', '=AVERAGEIF(A1:A5,1)', '=AVERAGEIFS(A1:A5,A1:A5,1)',
  '=WEEKNUM(A1,2)', '=GCD(A1:A5,12)', '=COUNTIFS(A1:A5,">1",A1:A5,"<5")', '=AVERAGEIF(A1:A5,">1",A1:A5)'
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $xl.Calculation = -4135   # xlCalculationManual
  $sheet = $wb.Worksheets.Item(1)
  $sheet.Name = 'Calc'
  1..5 | ForEach-Object { $sheet.Cells.Item($_, 1).Value2 = $_ }
  $row = 0
  foreach ($formula in $formulas) {
    $row++
    $sheet.Cells.Item($row, 3).Formula = $formula
  }
  $wb.SaveAs($xlsb, 50)   # xlExcel12
  $wb.SaveAs($xlsx, 51)   # xlOpenXMLWorkbook
  $wb.Close($false)
} finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Get-Item $xlsb, $xlsx | Select-Object Name, Length
