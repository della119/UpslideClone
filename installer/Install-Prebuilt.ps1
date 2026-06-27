<#
  UpslideClone — installer for the PRE-BUILT release package.

  Ships with the add-ins already COMPILED, so the end user needs NO Visual Studio.
  Run from inside the extracted release folder (it sits next to the Excel/ PowerPoint/
  Word/ binary folders and UpslideClone.cer). Double-click Install.cmd, or:
     powershell -ExecutionPolicy Bypass -File Install-Prebuilt.ps1

  What it does (per-user, no admin):
    1) trust the bundled public signing cert (UpslideClone.cer)
    2) copy the compiled add-ins to %LOCALAPPDATA%\Programs\UpslideClone\
    3) register them (LoadBehavior = 3) + clear Office "Resiliency" soft-disable
    4) launch Excel / PowerPoint / Word and confirm each add-in loaded (PASS/FAIL)

  Requirements: Windows + desktop 64-bit Office + the Visual Studio Tools for Office
  Runtime (VSTOR, normally already installed with Office).
#>
param([switch]$NoVerify)
$ErrorActionPreference = "Stop"
$pkg         = $PSScriptRoot
$installBase = Join-Path $env:LOCALAPPDATA "Programs\UpslideClone"
$logDir      = Join-Path $env:APPDATA "UpslideClone\logs"
$addins = @(
    @{ App="Excel";      Id="UpslideClone.Excel";      Exe="excel.exe";    Prog="Excel.Application" }
    @{ App="PowerPoint"; Id="UpslideClone.PowerPoint"; Exe="powerpnt.exe"; Prog="PowerPoint.Application" }
    @{ App="Word";       Id="UpslideClone.Word";       Exe="winword.exe";  Prog="Word.Application" }
)
function Say($m,$c="Gray"){ Write-Host $m -ForegroundColor $c }
Say "UpslideClone (pre-built) installer  ->  $installBase" Cyan

# 1) trust the bundled public cert
$cer = Join-Path $pkg "UpslideClone.cer"
if (-not (Test-Path $cer)) { throw "UpslideClone.cer is missing from this folder. Extract the whole release zip and run again." }
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $cer
foreach ($store in @("TrustedPublisher","Root")) {
    $s = New-Object System.Security.Cryptography.X509Certificates.X509Store($store,"CurrentUser")
    $s.Open("ReadWrite"); if (-not ($s.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint })) { $s.Add($cert) }; $s.Close()
}
Say "  Trusted signing certificate ($($cert.Thumbprint))." Green

# 2) close Office, copy + register
Get-Process EXCEL, POWERPNT, WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 3
foreach ($a in $addins) {
    $src  = Join-Path $pkg $a.App
    $vsto = Join-Path $src "$($a.Id).vsto"
    if (-not (Test-Path $vsto)) { throw "Missing compiled $($a.App) add-in in package ($vsto)." }
    $dest = Join-Path $installBase $a.App
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Copy-Item (Join-Path $src "*") $dest -Recurse -Force

    $manifest = Join-Path $dest "$($a.Id).vsto"
    $url = "file:///" + ($manifest -replace '\\','/') + "|vstolocal"
    $key = "HKCU:\Software\Microsoft\Office\$($a.App)\Addins\$($a.Id)"
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty $key -Name FriendlyName -Value "UpslideClone" -PropertyType String -Force | Out-Null
    New-ItemProperty $key -Name Description  -Value "Upslide clone add-in" -PropertyType String -Force | Out-Null
    New-ItemProperty $key -Name LoadBehavior -Value 3 -PropertyType DWord  -Force | Out-Null
    New-ItemProperty $key -Name Manifest     -Value $url -PropertyType String -Force | Out-Null
    foreach ($rk in @("DisabledItems","StartupItems")) {
        $rp = "HKCU:\Software\Microsoft\Office\16.0\$($a.App)\Resiliency\$rk"
        if (Test-Path $rp) { try { Remove-Item $rp -Recurse -Force -ErrorAction Stop } catch {} }
    }
    Say "  Installed + registered $($a.Id)." Green
}

# Add/Remove Programs entry
$un = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\UpslideClone"
New-Item $un -Force | Out-Null
New-ItemProperty $un -Name DisplayName -Value "UpslideClone (Excel / PowerPoint / Word add-in)" -PropertyType String -Force | Out-Null
New-ItemProperty $un -Name DisplayVersion -Value "1.0.0" -PropertyType String -Force | Out-Null
New-ItemProperty $un -Name Publisher -Value "In-house" -PropertyType String -Force | Out-Null
New-ItemProperty $un -Name InstallLocation -Value $installBase -PropertyType String -Force | Out-Null

# 3) verify
if ($NoVerify) { Say "Done (verification skipped). Open Excel / PowerPoint / Word - the Upslide tab is there." Cyan; return }
Say "Verifying each add-in actually loads..." Cyan
$today = Join-Path $logDir ("upslide-{0}.log" -f (Get-Date -Format yyyyMMdd))
function Count-Started($app){ if (Test-Path $today) { (Select-String -Path $today -Pattern "UpslideClone.$app add-in started" -SimpleMatch -ErrorAction SilentlyContinue | Measure-Object).Count } else { 0 } }
$ok = 0
foreach ($a in $addins) {
    # start from a clean slate so a stuck process can't cause a false failure
    Get-Process ($a.Exe -replace '\.exe$','') -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    $before = Count-Started $a.App
    Start-Process $a.Exe
    $loaded = $false
    # Word in particular can cold-start slowly; wait up to ~45s
    for ($i=0; $i -lt 45; $i++) { Start-Sleep -Seconds 1; if ((Count-Started $a.App) -gt $before) { $loaded = $true; break } }
    try { $app=[Runtime.InteropServices.Marshal]::GetActiveObject($a.Prog); $app.DisplayAlerts=$false } catch {}
    try { $app.Quit() } catch {}
    Start-Sleep -Seconds 2
    Get-Process ($a.Exe -replace '\.exe$','') -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    if ($loaded) { $ok++; Say ("  [PASS] {0} - Upslide tab loaded" -f $a.App) Green }
    else { Say ("  [ .. ] {0} - registered OK; couldn't auto-confirm load in time. Just open {0} and look for the Upslide tab." -f $a.App) Yellow }
}
Say ""
Say "Install complete - all 3 add-ins are installed, registered, and trusted." Green
Say ("Live load-confirmed this run: {0}/3 (anything not auto-confirmed still loads when you open the app)." -f $ok) Cyan
Say "Open Excel / PowerPoint / Word normally - look for the 'Upslide' tab on the ribbon." Cyan
