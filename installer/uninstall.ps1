<#
  UpslideClone — per-user uninstaller (W5). Removes the add-in registrations.
  Leaves the dev certificate in place (harmless); pass -RemoveCert to also remove it.

  Usage:  powershell -ExecutionPolicy Bypass -File uninstall.ps1 [-RemoveCert]
#>
param([switch]$RemoveCert)

$ErrorActionPreference = "Continue"

$addins = @(
    @{ App = "Excel";      Id = "UpslideClone.Excel" },
    @{ App = "PowerPoint"; Id = "UpslideClone.PowerPoint" },
    @{ App = "Word";       Id = "UpslideClone.Word" }
)

foreach ($a in $addins) {
    $key = "HKCU:\Software\Microsoft\Office\$($a.App)\Addins\$($a.Id)"
    if (Test-Path $key) { Remove-Item $key -Recurse -Force; Write-Host "Unregistered $($a.Id)." -ForegroundColor Green }
    else { Write-Host "$($a.Id) was not registered." }
}

if ($RemoveCert) {
    foreach ($store in @("TrustedPublisher", "Root")) {
        Get-ChildItem "Cert:\CurrentUser\$store" |
            Where-Object { $_.Subject -eq "CN=UpslideClone Dev" } |
            ForEach-Object { Remove-Item $_.PSPath -Force; Write-Host "Removed dev cert from $store." -ForegroundColor Green }
    }
}

Write-Host "Done. Restart Office apps to unload the add-in." -ForegroundColor Cyan
