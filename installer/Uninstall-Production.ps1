<#
  UpslideClone — production uninstall. Removes the registrations, the installed
  files, and the Add/Remove Programs entry. Leaves the dev cert (pass -RemoveCert
  to also remove it).
#>
param([switch]$RemoveCert)
$ErrorActionPreference = "Continue"

Get-Process EXCEL, POWERPNT, WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

foreach ($pair in @(@("Excel", "UpslideClone.Excel"), @("PowerPoint", "UpslideClone.PowerPoint"), @("Word", "UpslideClone.Word"))) {
    $key = "HKCU:\Software\Microsoft\Office\$($pair[0])\Addins\$($pair[1])"
    if (Test-Path $key) { Remove-Item $key -Recurse -Force; Write-Host "Unregistered $($pair[1])." -ForegroundColor Green }
}

$installBase = Join-Path $env:LOCALAPPDATA "Programs\UpslideClone"
if (Test-Path $installBase) { Remove-Item $installBase -Recurse -Force; Write-Host "Removed $installBase." -ForegroundColor Green }

$un = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\UpslideClone"
if (Test-Path $un) { Remove-Item $un -Recurse -Force; Write-Host "Removed Add/Remove Programs entry." -ForegroundColor Green }

if ($RemoveCert) {
    foreach ($store in @("TrustedPublisher", "Root")) {
        Get-ChildItem "Cert:\CurrentUser\$store" | Where-Object { $_.Subject -eq "CN=UpslideClone Dev" } |
            ForEach-Object { Remove-Item $_.PSPath -Force; Write-Host "Removed dev cert from $store." -ForegroundColor Green }
    }
}
Write-Host "Done. Restart Office apps to unload." -ForegroundColor Cyan
