$ErrorActionPreference = "Stop"

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$htmlPath = Join-Path $baseDir "F5_AppWorld_Seoul_2026_EDM.html"
$oftPath = Join-Path $baseDir "F5_AppWorld_Seoul_2026_EDM.oft"

if (-not (Test-Path $htmlPath)) {
  throw "HTML file not found: $htmlPath"
}

$cidMap = @{
  "assets/0430-edm-01-top.png" = "0430-edm-01-top.f5-appworld-seoul-2026"
  "assets/0430-edm-02-button-left.png" = "0430-edm-02-button-left.f5-appworld-seoul-2026"
  "assets/0430-edm-03-button.png" = "0430-edm-03-button.f5-appworld-seoul-2026"
  "assets/0430-edm-04-button-right.png" = "0430-edm-04-button-right.f5-appworld-seoul-2026"
  "assets/0430-edm-05-bottom.png" = "0430-edm-05-bottom.f5-appworld-seoul-2026"
}

$html = Get-Content -Path $htmlPath -Raw -Encoding UTF8
foreach ($relPath in $cidMap.Keys) {
  $html = $html.Replace("src=`"$relPath`"", "src=`"cid:$($cidMap[$relPath])`"")
}

$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.Subject = "F5 AppWorld Seoul 2026"

foreach ($relPath in $cidMap.Keys) {
  $imagePath = Join-Path $baseDir $relPath
  if (-not (Test-Path $imagePath)) {
    throw "Image file not found: $imagePath"
  }

  $attachment = $mail.Attachments.Add($imagePath)
  $propertyAccessor = $attachment.PropertyAccessor
  $propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", $cidMap[$relPath])
  $propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001F", "image/png")
}

$mail.HTMLBody = $html
$mail.SaveAs($oftPath, 2)

Write-Host "Created Outlook template: $oftPath"
