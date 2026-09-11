# Authors this case's fixture with real Excel: two pictures of one PNG on Sheet1. The first carries
# alternative text, a title, a crop on three edges and a link to a URL with a screen tip; the second
# links to a place in the workbook. Excel then saves `pictures.xlsx` in its own spelling of each.
#
# Requires Windows + desktop Excel. Not run by CI; re-run by hand only when the fixture must change:
#   pwsh -File test/corpus/fixtures/picture-properties-survive-roundtrip/author.ps1

$ErrorActionPreference = 'Stop'
$outDir = $PSScriptRoot
$xl = New-Object -ComObject Excel.Application
$png = Join-Path ([IO.Path]::GetTempPath()) 'picture-properties-author.png'
try {
  if ($xl.Workbooks.Count -ne 0) { $xl = $null; throw 'attached to a live Excel; refusing' }
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  Add-Type -AssemblyName System.Drawing
  $bmp = New-Object System.Drawing.Bitmap 40, 20
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.Clear([System.Drawing.Color]::SteelBlue)
  $g.Dispose()
  $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose()

  $wb = $xl.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $logo = $ws.Shapes.AddPicture($png, 0, -1, 10, 10, 120, 60)
  $logo.AlternativeText = 'Company logo'
  $logo.Title = 'Logo'
  $logo.PictureFormat.CropLeft = 3
  $logo.PictureFormat.CropRight = 6
  $logo.PictureFormat.CropTop = 1.5
  [void]$ws.Hyperlinks.Add($logo, 'https://example.com/about', '', 'About us')
  $jump = $ws.Shapes.AddPicture($png, 0, -1, 10, 100, 60, 30)
  [void]$ws.Hyperlinks.Add($jump, '', 'Sheet1!C3', 'Jump')

  $path = Join-Path $outDir 'pictures.xlsx'
  if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
  $wb.SaveAs($path, 51)
  $wb.Close($false)
} finally {
  if ($xl) { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
  [GC]::Collect()
  Remove-Item -LiteralPath $png -Force -ErrorAction SilentlyContinue
}
