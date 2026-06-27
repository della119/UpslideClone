<#
  UpslideClone — uninstaller for the pre-built release. Per-user, no admin.
  Removes the registry registrations, the installed files, and the Add/Remove entry.
  (Leaves the trusted certificate and your logs in place.)
#>
$ErrorActionPreference = "SilentlyContinue"
$installBase = Join-Path $env:LOCALAPPDATA "Programs\UpslideClone"
Get-Process EXCEL, POWERPNT, WINWORD | Stop-Process -Force
Start-Sleep -Seconds 2
foreach ($app in "Excel","PowerPoint","Word") {
    Remove-Item "HKCU:\Software\Microsoft\Office\$app\Addins\UpslideClone.$app" -Recurse -Force
}
Remove-Item $installBase -Recurse -Force
Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\UpslideClone" -Recurse -Force
Write-Host "UpslideClone uninstalled." -ForegroundColor Cyan
