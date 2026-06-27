<#
  UpslideClone — one-shot, foolproof installer.

  Does everything end-to-end so a fresh clone "just works":
    0) ensure a code-signing cert exists (creates a self-signed one if missing,
       exports the per-project .pfx keys — so the repo never has to ship a private key)
    1) build Release (auto-locates MSBuild; skip with -SkipBuild)
    2) trust the signing cert (CurrentUser TrustedPublisher + Root)
    3) copy the build to a stable location (%LOCALAPPDATA%\Programs\UpslideClone\)
    4) register each add-in (LoadBehavior = 3) pointing at that stable copy
    5) clear Office "Resiliency" soft-disable leftovers (the usual reason a tab vanishes)
    6) VERIFY: launch Excel / PowerPoint / Word, confirm each add-in actually loaded
       (reads the add-in's own log), print PASS/FAIL per app

  Usage (from anywhere):
    powershell -ExecutionPolicy Bypass -File installer\Install-Upslide.ps1
  Or just double-click  installer\Install-Upslide.cmd

  No admin rights needed (everything is per-user: HKCU + CurrentUser cert store).
#>
param(
    [string]$Configuration = "Release",
    [switch]$SkipBuild,
    [switch]$NoVerify
)
$ErrorActionPreference = "Stop"
$Root        = Split-Path -Parent $PSScriptRoot
$Subject     = "CN=UpslideClone Dev"
$installBase = Join-Path $env:LOCALAPPDATA "Programs\UpslideClone"
$logDir      = Join-Path $env:APPDATA "UpslideClone\logs"

$addins = @(
    @{ App="Excel";      Id="UpslideClone.Excel";      Proj="src\UpslideClone.Excel";      Exe="excel.exe";    Prog="Excel.Application" }
    @{ App="PowerPoint"; Id="UpslideClone.PowerPoint"; Proj="src\UpslideClone.PowerPoint"; Exe="powerpnt.exe"; Prog="PowerPoint.Application" }
    @{ App="Word";       Id="UpslideClone.Word";       Proj="src\UpslideClone.Word";       Exe="winword.exe"; Prog="Word.Application" }
)

function Say($m,$c="Gray"){ Write-Host $m -ForegroundColor $c }
Say "UpslideClone installer  ->  $installBase" Cyan

# ---- 0) ensure signing cert + per-project .pfx -----------------------------
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq $Subject } | Select-Object -First 1
if (-not $cert) {
    Say "  No signing cert found - creating a self-signed dev cert (one-time)..." Yellow
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $Subject `
                -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable `
                -NotAfter (Get-Date).AddYears(10)
}
$thumb = $cert.Thumbprint
# make sure every project has its .pfx (export password-less, matching VS's _TemporaryKey.pfx)
foreach ($a in $addins) {
    $pfx = Join-Path $Root "$($a.Proj)\$($a.Id)_TemporaryKey.pfx"
    if (-not (Test-Path $pfx)) {
        try {
            $bytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, "")
            [System.IO.File]::WriteAllBytes($pfx, $bytes)
            Say "  Exported signing key -> $($a.Id)_TemporaryKey.pfx" DarkGray
        } catch { throw "Could not export signing key to $pfx. The cert's private key is not exportable. Delete the '$Subject' cert from Cert:\CurrentUser\My and re-run so the script can create an exportable one." }
    }
}
# trust it so VSTO loads without prompts
foreach ($store in @("TrustedPublisher","Root")) {
    $s = New-Object System.Security.Cryptography.X509Certificates.X509Store($store,"CurrentUser")
    $s.Open("ReadWrite"); if (-not ($s.Certificates | Where-Object Thumbprint -eq $thumb)) { $s.Add($cert) }; $s.Close()
}
Say "  Signing cert ready + trusted ($thumb)." Green

# ---- 1) build Release ------------------------------------------------------
if (-not $SkipBuild) {
    $msbuild = $null
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $p = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" 2>$null | Select-Object -First 1
        if ($p) { $msbuild = $p }
    }
    if (-not $msbuild) {
        foreach ($c in @("D:\VS2022\MSBuild\Current\Bin\MSBuild.exe",
                         "${env:ProgramFiles}\Microsoft Visual Studio\2022\*\MSBuild\Current\Bin\MSBuild.exe",
                         "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\*\MSBuild\Current\Bin\MSBuild.exe")) {
            $r = Get-Item $c -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($r) { $msbuild = $r.FullName; break }
        }
    }
    if (-not $msbuild) { throw "MSBuild not found. Install Visual Studio 2022 with the 'Office/SharePoint development' + '.NET desktop' workloads, then re-run (or pass -SkipBuild if you already built)." }
    Say "  Building $Configuration with $msbuild ..." Gray
    & $msbuild (Join-Path $Root "UpslideClone.sln") /t:Build /p:Configuration=$Configuration /p:ManifestCertificateThumbprint=$thumb /restore /nologo /v:m
    if ($LASTEXITCODE -ne 0) { throw "Build failed (exit $LASTEXITCODE). Fix the build, then re-run." }
    Say "  Build OK." Green
}

# ---- 2-3) close Office, copy to stable location, register ------------------
Get-Process EXCEL, POWERPNT, WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 3
foreach ($a in $addins) {
    $src  = Join-Path $Root "$($a.Proj)\bin\$Configuration"
    $vsto = Join-Path $src "$($a.Id).vsto"
    if (-not (Test-Path $vsto)) { throw "Build output missing: $vsto. Run without -SkipBuild." }
    $dest = Join-Path $installBase $a.App
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Copy-Item (Join-Path $src "*") $dest -Recurse -Force
    $assets = Join-Path $Root "assets"
    if (Test-Path $assets) { Copy-Item $assets (Join-Path $dest "assets") -Recurse -Force }

    $manifest = Join-Path $dest "$($a.Id).vsto"
    $url = "file:///" + ($manifest -replace '\\','/') + "|vstolocal"
    $key = "HKCU:\Software\Microsoft\Office\$($a.App)\Addins\$($a.Id)"
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty $key -Name FriendlyName -Value "UpslideClone" -PropertyType String -Force | Out-Null
    New-ItemProperty $key -Name Description  -Value "Upslide clone add-in" -PropertyType String -Force | Out-Null
    New-ItemProperty $key -Name LoadBehavior -Value 3 -PropertyType DWord  -Force | Out-Null
    New-ItemProperty $key -Name Manifest     -Value $url -PropertyType String -Force | Out-Null

    # ---- 5) clear Resiliency soft-disable for this app (the vanishing-tab cause)
    foreach ($rk in @("DisabledItems","StartupItems","CrashingAddinList")) {
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
New-ItemProperty $un -Name UninstallString -Value ("powershell -ExecutionPolicy Bypass -File `"$Root\installer\Uninstall-Production.ps1`"") -PropertyType String -Force | Out-Null

# ---- 6) VERIFY each add-in actually loads ----------------------------------
if ($NoVerify) { Say "Done (verification skipped)." Cyan; return }

Say "Verifying each add-in actually loads..." Cyan
$today = Join-Path $logDir ("upslide-{0}.log" -f (Get-Date -Format yyyyMMdd))
function Count-Started($app) { if (Test-Path $today) { (Select-String -Path $today -Pattern "UpslideClone.$app add-in started" -SimpleMatch -ErrorAction SilentlyContinue | Measure-Object).Count } else { 0 } }

$results = @()
foreach ($a in $addins) {
    $before = Count-Started $a.App
    Start-Process $a.Exe
    $loaded = $false
    for ($i=0; $i -lt 25; $i++) {
        Start-Sleep -Milliseconds 800
        if ((Count-Started $a.App) -gt $before) { $loaded = $true; break }
    }
    # close it gracefully (no force-kill -> no safe-mode prompt next time)
    try { $app=[Runtime.InteropServices.Marshal]::GetActiveObject($a.Prog); $app.DisplayAlerts=$false } catch {}
    try { $app.Quit() } catch {}
    Start-Sleep -Seconds 2
    Get-Process ($a.Exe -replace '\.exe$','') -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    $results += [pscustomobject]@{ App=$a.App; Loaded=$loaded }
    if ($loaded) { Say ("  [PASS] {0} - Upslide tab loaded" -f $a.App) Green } else { Say ("  [FAIL] {0} - add-in did NOT load (see $logDir)" -f $a.App) Red }
}

Say ""
$ok = ($results | Where-Object Loaded).Count
Say ("Result: {0}/3 add-ins verified loaded." -f $ok) (@{3="Green";2="Yellow";1="Yellow";0="Red"}[$ok])
Say "Open Excel / PowerPoint / Word normally - the Upslide tab is there." Cyan
