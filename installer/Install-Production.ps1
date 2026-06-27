<#
  UpslideClone — production per-user install.

  Unlike the dev installer (which points Office at the churning bin\Debug folder),
  this COPIES the Release build to a stable location
  (%LOCALAPPDATA%\Programs\UpslideClone\) and registers the add-ins there. Because
  the registered files never change underneath Office, the Upslide tab loads on a
  normal app launch — no cold-start launcher needed. Also adds an Add/Remove
  Programs entry. No admin rights required (HKCU + CurrentUser cert store).

  Usage:  powershell -ExecutionPolicy Bypass -File installer\Install-Production.ps1
          (build Release first:  msbuild UpslideClone.sln /t:Build /p:Configuration=Release /restore)
#>
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$Configuration = "Release"
)
$ErrorActionPreference = "Stop"
$installBase = Join-Path $env:LOCALAPPDATA "Programs\UpslideClone"
Write-Host "UpslideClone production install -> $installBase" -ForegroundColor Cyan

$addins = @(
    @{ App = "Excel";      Id = "UpslideClone.Excel";      Proj = "src\UpslideClone.Excel" },
    @{ App = "PowerPoint"; Id = "UpslideClone.PowerPoint"; Proj = "src\UpslideClone.PowerPoint" },
    @{ App = "Word";       Id = "UpslideClone.Word";       Proj = "src\UpslideClone.Word" }
)

# Close any running Office so files aren't locked.
Get-Process EXCEL, POWERPNT, WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 3

# 1) Trust the signing cert.
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=UpslideClone Dev" } | Select-Object -First 1
if (-not $cert) { throw "Signing cert 'CN=UpslideClone Dev' not found in CurrentUser\My. Build once in Visual Studio, or create it (see installer\README.md)." }
foreach ($store in @("TrustedPublisher", "Root")) {
    $s = New-Object System.Security.Cryptography.X509Certificates.X509Store($store, "CurrentUser")
    $s.Open("ReadWrite"); $s.Add($cert); $s.Close()
}
Write-Host "  Trusted signing certificate ($($cert.Thumbprint))." -ForegroundColor Green

# 2) Copy each add-in's Release output to the stable location + register it there.
foreach ($a in $addins) {
    $src = Join-Path $Root "$($a.Proj)\bin\$Configuration"
    $vsto = Join-Path $src "$($a.Id).vsto"
    if (-not (Test-Path $vsto)) { throw "Release build not found: $vsto. Build Release first." }

    $dest = Join-Path $installBase $a.App
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Copy-Item (Join-Path $src "*") $dest -Recurse -Force
    # bundle the editable theme next to the assembly (optional; falls back to built-in)
    $assets = Join-Path $Root "assets"
    if (Test-Path $assets) { Copy-Item $assets (Join-Path $dest "assets") -Recurse -Force }

    $manifest = Join-Path $dest "$($a.Id).vsto"
    $url = "file:///" + ($manifest -replace '\\', '/') + "|vstolocal"
    $key = "HKCU:\Software\Microsoft\Office\$($a.App)\Addins\$($a.Id)"
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name FriendlyName -Value "UpslideClone" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name Description  -Value "Upslide clone add-in" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $key -Name Manifest -Value $url -PropertyType String -Force | Out-Null
    Write-Host "  Installed + registered $($a.Id) under $($a.App)." -ForegroundColor Green
}

# 3) Add/Remove Programs entry.
$un = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\UpslideClone"
New-Item -Path $un -Force | Out-Null
New-ItemProperty -Path $un -Name DisplayName -Value "UpslideClone (Excel / PowerPoint / Word add-in)" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $un -Name DisplayVersion -Value "1.0.0" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $un -Name Publisher -Value "In-house" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $un -Name InstallLocation -Value $installBase -PropertyType String -Force | Out-Null
New-ItemProperty -Path $un -Name UninstallString -Value ("powershell -ExecutionPolicy Bypass -File `"$Root\installer\Uninstall-Production.ps1`"") -PropertyType String -Force | Out-Null
New-ItemProperty -Path $un -Name NoModify -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $un -Name NoRepair -Value 1 -PropertyType DWord -Force | Out-Null

Write-Host "Done. Open Excel / PowerPoint / Word normally — the Upslide tab loads from the stable install." -ForegroundColor Cyan
