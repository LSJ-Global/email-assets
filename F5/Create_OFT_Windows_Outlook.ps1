$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$htmlPath = Join-Path $baseDir "F5_AppWorld_Seoul_2026_EDM.html"
$oftPath = Join-Path $baseDir "F5_AppWorld_Seoul_2026_EDM.oft"

if (-not (Test-Path -LiteralPath $htmlPath)) {
  throw "HTML file not found: $htmlPath"
}

$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.Subject = "F5 AppWorld Seoul 2026"
$mail.BodyFormat = 2

# The HTML already points to hosted raw GitHub image URLs.
$html = Get-Content -LiteralPath $htmlPath -Raw -Encoding UTF8

$mail.HTMLBody = $html
$mail.SaveAs($oftPath, 2)

Write-Host "Created OFT: $oftPath"
Write-Host "Images are loaded from GitHub raw URLs. If Outlook blocks external images, click Download Pictures or allow automatic image downloads for this sender/domain."
