# UpslideClone — Deployment

Per-user install of the three VSTO add-ins (Excel, PowerPoint, Word). No admin
rights required — everything is under `HKCU` and the CurrentUser certificate stores.

## ⭐ One-click install (recommended) — `Install-Upslide.cmd`
**Just double-click `installer\Install-Upslide.cmd`.** One script does everything and
*verifies it worked*:

0. ensures a code-signing cert exists — **creates a self-signed one if missing** and
   exports the per-project `.pfx` keys, so a fresh `git clone` works with **no private
   key committed to the repo**
1. builds **Release** (auto-locates MSBuild via `vswhere`)
2. trusts the signing cert (CurrentUser TrustedPublisher + Root)
3. copies the build to a stable location (`%LOCALAPPDATA%\Programs\UpslideClone\`)
4. registers each add-in (`LoadBehavior = 3`) pointing at that stable copy
5. clears Office **Resiliency** soft-disable leftovers (the usual "tab vanished" cause)
6. **launches Excel / PowerPoint / Word and confirms each add-in actually loaded**,
   printing `[PASS]/[FAIL]` per app

No admin rights needed (per-user: HKCU + CurrentUser cert store). Equivalent PowerShell:
```powershell
powershell -ExecutionPolicy Bypass -File installer\Install-Upslide.ps1
```
Flags: `-SkipBuild` (re-register without rebuilding), `-NoVerify` (skip the launch test).

### If the tab ever disappears → `Repair-Upslide.cmd`
Double-click it. Re-registers + clears the soft-disable + re-verifies, **without** a
rebuild (≈30 s). It's `Install-Upslide.ps1 -SkipBuild` under the hood.

---

## Legacy: `Install-Production.ps1` (manual build first)
Older two-step path — build Release yourself, then register. Superseded by
`Install-Upslide.cmd` above (which also auto-builds, clears Resiliency, and verifies).
```powershell
# 1. Build Release
& "D:\VS2022\MSBuild\Current\Bin\MSBuild.exe" UpslideClone.sln /t:Build /p:Configuration=Release /restore

# 2. Install for the current user
powershell -ExecutionPolicy Bypass -File installer\Install-Production.ps1
```
Then just **open Excel / PowerPoint / Word normally** — the Upslide tab is there. No
cold-start launcher needed. To remove: `installer\Uninstall-Production.ps1` (or via
Settings ▸ Apps).

> Why this fixes the "disappearing tab": the dev install below points Office at the
> `bin\Debug` folder, which changes on every rebuild and goes stale. The production
> install points at a fixed copy that never moves underneath Office.

## Dev install (for active development against `bin\Debug`)
```powershell
& "D:\VS2022\MSBuild\Current\Bin\MSBuild.exe" UpslideClone.sln /t:Build /p:Configuration=Debug /restore
powershell -ExecutionPolicy Bypass -File installer\install.ps1
```
With the dev install, open via `tools\Launch-Excel.cmd` (etc.) and **re-run `install.ps1`
after every rebuild** (the signed manifests change, so the registration must be re-synced).

## Uninstall
```powershell
powershell -ExecutionPolicy Bypass -File installer\uninstall.ps1            # remove registrations
powershell -ExecutionPolicy Bypass -File installer\uninstall.ps1 -RemoveCert # also remove the dev cert
```

## Notes
- `install.ps1` finds the dev signing cert (`CN=UpslideClone Dev`) in the CurrentUser\My
  store, trusts it in CurrentUser **TrustedPublisher** + **Root**, then registers each
  add-in under `HKCU\Software\Microsoft\Office\<App>\Addins\` with a `|vstolocal` manifest
  URL pointing at `bin\<Config>`. (The cert is created once when you first build/sign in
  Visual Studio, or via `New-SelfSignedCertificate -Type CodeSigningCert -Subject
  "CN=UpslideClone Dev" -CertStoreLocation Cert:\CurrentUser\My`.)
- The PowerPoint **Content Library** stores items under `Documents\UpslideClone\` (Office
  refuses to save into `%APPDATA%`). Logs are at `%APPDATA%\UpslideClone\logs`.
- For a **machine-wide, production** install, replace the dev self-signed cert with
  a real code-signing certificate and switch to a ClickOnce publish or a WiX/MSI
  (registers under `HKLM` + Trusted Publishers). The dev cert here is for internal
  testing only.
- `-Configuration Release` installs the Release build instead of Debug.
