<#
  UpslideClone — per-user installer (W5).
  Registers the Excel, PowerPoint and Word VSTO add-ins for the current user and
  trusts the signing certificate so Office loads them. Run from an elevated prompt
  is NOT required (HKCU + CurrentUser cert stores only).

  Usage:   powershell -ExecutionPolicy Bypass -File install.ps1 [-Root <repoRoot>] [-Configuration Debug|Release]
#>
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"
Write-Host "UpslideClone installer - root: $Root ($Configuration)" -ForegroundColor Cyan

$addins = @(
    @{ App = "Excel";      Id = "UpslideClone.Excel";      Proj = "src\UpslideClone.Excel" },
    @{ App = "PowerPoint"; Id = "UpslideClone.PowerPoint"; Proj = "src\UpslideClone.PowerPoint" },
    @{ App = "Word";       Id = "UpslideClone.Word";       Proj = "src\UpslideClone.Word" }
)

# 1) Trust the manifest-signing certificate in CurrentUser. VSTO signs the
#    ClickOnce manifests with a cert from CurrentUser\My; trust it as a publisher.
$dll = Join-Path $Root "$($addins[0].Proj)\bin\$Configuration\$($addins[0].Id).dll"
if (-not (Test-Path $dll)) {
    throw "Build output not found: $dll. Build the solution first (msbuild UpslideClone.sln /restore)."
}
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=UpslideClone Dev" } | Select-Object -First 1
if (-not $cert) {
    throw "Signing cert 'CN=UpslideClone Dev' not found in CurrentUser\My. Build in Visual Studio once, or import the project's *_TemporaryKey.pfx."
}
foreach ($store in @("TrustedPublisher", "Root")) {
    $s = New-Object System.Security.Cryptography.X509Certificates.X509Store($store, "CurrentUser")
    $s.Open("ReadWrite"); $s.Add($cert); $s.Close()
}
Write-Host "  Trusted signing certificate ($($cert.Thumbprint))." -ForegroundColor Green

# 2) Register each add-in (HKCU) pointing at its built .vsto manifest.
foreach ($a in $addins) {
    $manifest = Join-Path $Root "$($a.Proj)\bin\$Configuration\$($a.Id).vsto"
    if (-not (Test-Path $manifest)) { Write-Warning "  Skipping $($a.Id): $manifest not found."; continue }
    $url = "file:///" + ($manifest -replace '\\','/') + "|vstolocal"
    $key = "HKCU:\Software\Microsoft\Office\$($a.App)\Addins\$($a.Id)"
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name FriendlyName -Value "UpslideClone" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name Description  -Value "Upslide clone add-in" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $key -Name Manifest -Value $url -PropertyType String -Force | Out-Null
    Write-Host "  Registered $($a.Id) under $($a.App)." -ForegroundColor Green
}

Write-Host "Done. Restart Excel / PowerPoint / Word to load the Upslide tab." -ForegroundColor Cyan
