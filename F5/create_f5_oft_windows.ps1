param(
    [string]$HtmlPath = (Join-Path $PSScriptRoot "f5_appworld_seoul_2026_vip_dinner.html"),
    [string]$OutputPath = (Join-Path $PSScriptRoot "f5_appworld_seoul_2026_vip_dinner.oft"),
    [string]$Subject = "F5 AppWorld Seoul 2026 VIP Dinner 초청"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $HtmlPath)) {
    throw "HTML file not found: $HtmlPath"
}

$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.Subject = $Subject
$mail.BodyFormat = 2
$mail.HTMLBody = Get-Content -LiteralPath $HtmlPath -Raw -Encoding UTF8

# 2 is Outlook.OlSaveAsType.olTemplate. This must be run on Windows with Outlook installed.
$mail.SaveAs($OutputPath, 2)
Write-Host "Created Outlook template: $OutputPath"
