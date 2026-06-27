<#
  UpslideClone — build the pre-built RELEASE package (a zip end users can install
  without Visual Studio).

  Produces:  dist\UpslideClone-v<Version>.zip
  containing: compiled Excel/PowerPoint/Word add-ins, the public signing cert,
              a one-click Install.cmd / Uninstall.cmd, and a README.txt.

  Usage:
    powershell -ExecutionPolicy Bypass -File installer\Make-Release.ps1            # build Release first
    powershell -ExecutionPolicy Bypass -File installer\Make-Release.ps1 -SkipBuild # use existing bin\Release
#>
param(
    [string]$Version = "1.0.1",
    [switch]$SkipBuild
)
$ErrorActionPreference = "Stop"
$Root      = Split-Path -Parent $PSScriptRoot
$Subject   = "CN=UpslideClone Dev"
$stageRoot = Join-Path $Root "dist"
$stage     = Join-Path $stageRoot "UpslideClone-v$Version"
$zipPath   = Join-Path $stageRoot "UpslideClone-v$Version.zip"
$addins = @(
    @{ App="Excel";      Id="UpslideClone.Excel";      Proj="src\UpslideClone.Excel" }
    @{ App="PowerPoint"; Id="UpslideClone.PowerPoint"; Proj="src\UpslideClone.PowerPoint" }
    @{ App="Word";       Id="UpslideClone.Word";       Proj="src\UpslideClone.Word" }
)
function Say($m,$c="Gray"){ Write-Host $m -ForegroundColor $c }
Say "Building release package v$Version -> $zipPath" Cyan

# 0) ensure a Release build exists (optionally build it)
if (-not $SkipBuild) {
    $msbuild = $null
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) { $msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" 2>$null | Select-Object -First 1 }
    if (-not $msbuild) { foreach ($c in @("D:\VS2022\MSBuild\Current\Bin\MSBuild.exe","${env:ProgramFiles}\Microsoft Visual Studio\2022\*\MSBuild\Current\Bin\MSBuild.exe")) { $r = Get-Item $c -ErrorAction SilentlyContinue | Select-Object -First 1; if ($r) { $msbuild = $r.FullName; break } } }
    if (-not $msbuild) { throw "MSBuild not found (need VS2022). Or pass -SkipBuild if bin\Release already exists." }
    & $msbuild (Join-Path $Root "UpslideClone.sln") /t:Build /p:Configuration=Release /restore /nologo /v:m
    if ($LASTEXITCODE -ne 0) { throw "Release build failed." }
}

# 1) fresh staging dir
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force -Path $stage | Out-Null

# 2) copy compiled add-ins (exclude private key + debug symbols)
foreach ($a in $addins) {
    $src = Join-Path $Root "$($a.Proj)\bin\Release"
    if (-not (Test-Path (Join-Path $src "$($a.Id).vsto"))) { throw "Release build missing for $($a.App). Build Release or drop -SkipBuild." }
    $dest = Join-Path $stage $a.App
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Get-ChildItem $src -Recurse -File | Where-Object { $_.Extension -notin ".pfx",".pdb" } | ForEach-Object {
        $rel = $_.FullName.Substring($src.Length).TrimStart('\')
        $td  = Join-Path $dest $rel
        New-Item -ItemType Directory -Force -Path (Split-Path $td) | Out-Null
        Copy-Item $_.FullName $td -Force
    }
}
# theme assets
$assets = Join-Path $Root "assets"
if (Test-Path $assets) { Copy-Item $assets (Join-Path $stage "assets") -Recurse -Force }

# 3) export the PUBLIC signing cert (no private key)
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq $Subject } | Select-Object -First 1
if (-not $cert) { throw "Signing cert '$Subject' not found. Run installer\Install-Upslide.ps1 once (it creates it), then re-run." }
[System.IO.File]::WriteAllBytes((Join-Path $stage "UpslideClone.cer"), $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))

# 4) installer scripts + launchers + readme
Copy-Item (Join-Path $PSScriptRoot "Install-Prebuilt.ps1")   (Join-Path $stage "Install-Prebuilt.ps1")   -Force
Copy-Item (Join-Path $PSScriptRoot "Uninstall-Prebuilt.ps1") (Join-Path $stage "Uninstall-Prebuilt.ps1") -Force

@"
@echo off
echo Installing UpslideClone add-ins (Excel / PowerPoint / Word)...
echo This trusts the bundled certificate, registers the add-ins, and verifies they load.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Prebuilt.ps1" %*
echo.
pause
"@ | Set-Content (Join-Path $stage "Install.cmd") -Encoding ASCII

@"
@echo off
echo Removing UpslideClone add-ins...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall-Prebuilt.ps1" %*
echo.
pause
"@ | Set-Content (Join-Path $stage "Uninstall.cmd") -Encoding ASCII

@"
UpslideClone $Version - in-house Office add-in suite (Excel / PowerPoint / Word)
================================================================================

WHAT IT IS
  Branded one-click formatting, finance charts (waterfall / CAGR), an Excel ->
  PowerPoint / Word live-link engine, and a PowerPoint slide-design toolkit.

REQUIREMENTS
  * Windows 10 / 11
  * Desktop Microsoft Office (Excel / PowerPoint / Word) - NOT the web or Mac versions
  * Visual Studio Tools for Office Runtime (VSTOR) - normally already installed with Office
  No Visual Studio and no admin rights needed - the add-ins in this package are
  already compiled.

INSTALL
  1. Extract this whole zip to a folder (keep all files together).
  2. Double-click  Install.cmd
  3. It trusts the bundled certificate, installs to
       %LOCALAPPDATA%\Programs\UpslideClone\
     registers the add-ins, then opens Excel / PowerPoint / Word to confirm each
     one loaded (you'll see [PASS]/[FAIL]).
  4. Open Excel / PowerPoint / Word - look for the "Upslide" tab on the ribbon.

IF THE TAB EVER DISAPPEARS
  Re-run Install.cmd (it re-registers and clears Office's safe-mode soft-disable).
  If Office ever asks to start in Safe Mode, click No.

UNINSTALL
  Double-click  Uninstall.cmd  (or remove "UpslideClone" from Settings > Apps).

CERTIFICATE NOTE (please read)
  The add-ins are signed with a SELF-SIGNED certificate (UpslideClone.cer, public
  key only). Install.cmd adds it to your CurrentUser Trusted Publishers and Root
  stores so Office will load the add-ins without a prompt - this is unavoidable for
  self-signed VSTO add-ins (Windows won't trust them otherwise). It is fine for
  INTERNAL / known-source use; only install if you trust the source. For wide public
  distribution, the add-ins should instead be signed with a CA-issued code-signing
  certificate, so no manual trust step is needed and the publisher is verifiable.

SOURCE
  https://github.com/della119/UpslideClone
"@ | Set-Content (Join-Path $stage "README.txt") -Encoding ASCII

# 5) zip it
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path (Join-Path $stage "*") -DestinationPath $zipPath -Force
$mb = "{0:N1}" -f ((Get-Item $zipPath).Length / 1MB)
Say "Package ready: $zipPath ($mb MB)" Green
Say "Contents:" Gray
Get-ChildItem $stage | Select-Object Name | Format-Table -HideTableHeaders | Out-String | Write-Host
