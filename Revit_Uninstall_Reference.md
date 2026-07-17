# Autodesk Revit Uninstall — Teardown Reference

Working knowledge for scripting a clean, scoped uninstall of an Autodesk Revit
product on Windows without disturbing shared components or other Autodesk apps.
Derived from a verified Revit 2026 removal on machine `ICECREAMASSASIN`
(2026-07-17). The delivered script is `Uninstall-Revit2026.ps1` (PowerShell 5.1
compatible, self-elevating). Published to GitHub with an MIT LICENSE and a
README (author: MrGezz).

## What worked

The core Revit 2026 product uninstalled with exit 0 via Autodesk's ODIS
installer. The MSI-based add-ins uninstalled instantly by product code.
Shared components (material libraries, licensing, RealDWG, Content Catalog)
were preserved, and AutoCAD/Navisworks 2026 (separate apps) were untouched.

## Five lessons that cost debugging cycles

1. **Never route an Autodesk uninstall command through `cmd /c`.** The ODIS
   `UninstallString` is *unquoted* and its executable path contains a space
   (`C:\Program Files\Autodesk\AdODIS\V1\installer.exe`). `cmd /c` reads the
   exe as `C:\Program` and fails instantly with a generic **exit 1**. Fix:
   parse the command line into executable + argument string and call
   `Start-Process -FilePath <exe> -ArgumentList <args>` directly — Start-Process
   quotes the path correctly. Split unquoted command lines on the first `.exe`
   token so spaced paths survive.

2. **Under `Set-StrictMode -Version Latest`, do not pipe an absent property to
   `ForEach-Object`.** `$obj.PSObject.Properties['X'] | ForEach-Object { $obj.X }`
   still runs the block once with `$null` when the value is absent, then throws
   `PropertyNotFoundStrict`. Use a guarded accessor instead:
   `$p = $obj.PSObject.Properties[$Name]; if ($p) { $p.Value } else { $null }`.

3. **Gate residual-file cleanup on uninstall success.** An early version deleted
   `C:\Program Files\Autodesk\Revit 2026` and the per-user settings folders even
   though the core uninstall had failed, leaving a half-removed state (files gone,
   registration still present). Residual cleanup must run only when
   `$failures -eq 0`.

4. **Do not `return ,$array` when the caller wraps the result in `@(...)`.** The
   unary-comma operator double-nests the array. `@()` then unrolls only the outer
   layer, so `foreach ($x in $result)` iterates **once** over the whole inner
   array; `$x.Prop` triggers member-access enumeration and returns an array of
   every element's property. Symptom seen here: two uninstall candidates fused
   into one, `Start-Process -FilePath` got `System.Object[]`, and threw "Cannot
   convert 'System.Object[]' to the type 'System.String'". Fix: plain
   `return $array` and let the caller's `@(...)` normalize (correct for 0/1/many).

5. **`-Force` (a custom switch) does NOT suppress PowerShell's ShouldProcess
   confirmation.** With `[CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]`,
   `ShouldProcess` still prompts "Are you sure?" for every item even when a custom
   `-Force` skipped your own `Read-Host`. Fix: when `-Force` is set, also set
   `$ConfirmPreference = 'None'` in script scope (inherited by nested advanced
   functions). Note MSI `/qn` is fully silent; ODIS EXE uninstallers may still
   show their own progress UI — `--silent` is a valid AdODIS 2026 flag (confirmed
   working, exit 0), but keep the exact vendor command as an automatic fallback
   so a wrong flag can't block removal.

## Product-selection rule (the reliable part)

- **Core:** display name matches `Autodesk Revit <year>`.
- **Orphaned add-ins/content (opt-in, default on):** any product whose display
  name contains **both** `Revit`/`RVT` **and** the target year. This single rule
  catches add-ins, content packs, exporters, DB Link, IFC, interop tools, etc.
  without an exhaustive per-product list.
- **Always exclude** (shared / cross-version — other Revit years depend on them):
  - Version-range packs, e.g. names matching `\d{4}\s*[-–]\s*\d{4}` (`2024-2027`).
  - `Content Catalog`, `RealDWG`, material libraries, licensing, Genuine Service,
    Identity Manager, Autodesk Access, ODIS/AdODIS, Desktop Connector, and any
    version-neutral interop manager (no year in the name → excluded automatically).
- **Why AutoCAD/Navisworks are safe:** their names contain no `Revit`, so the
  "Revit + year" rule never matches them even when they share the year.
- **RealDWG is not AutoCAD's.** `Autodesk RealDWG Shared <year>` is the
  redistributable DWG read/write engine for *non-AutoCAD* apps (Revit, Navisworks,
  Civil 3D...). AutoCAD uses its own core and does not need it. Keep it while any
  non-AutoCAD DWG consumer of that year remains (e.g. Navisworks 2026).
- **Note:** a product GUID may coincidentally contain the year digits (e.g.
  `{05BC6921-2026-49D7-...}`); this is harmless because matching is by display
  name, not by GUID. De-dupe msiexec candidates by extracted GUID (case- and
  flag-order-insensitive) so an MSI-code candidate and an msiexec
  `UninstallString` for the same product don't both run.

## Uninstall command resolution order (per product)

1. MSI product code — `msiexec.exe /x {GUID} /qn /norestart` — when
   `WindowsInstaller = 1` and the registry key name is a GUID. Most deterministic.
2. `QuietUninstallString` (run the exe directly, not via cmd).
3. Raw `UninstallString` (run the exe directly; coerce msiexec `/I`→`/X` and add
   `/qn /norestart`; for ODIS EXE uninstallers, try `--silent` first with the
   exact vendor command as fallback). This is the ODIS path for the core product.

Treat exit codes `0`, `3010` (reboot required), and `1605` (already gone) as
success. Try each candidate in order; fall through on any other code.

## Captured identifiers (Revit 2026 on this machine)

- Core Revit 2026 version: `26.4.20.9`
- Core Revit 2026 ODIS bundle GUID (metadata folder):
  `{8986CA21-EA9C-32F3-A1DB-C34BD2BDA7A5}`
- Batch Print for Revit 2026 (MSI): `{82AF00E4-2601-0010-0000-FCE0F8702600}`
- eTransmit for Revit 2026 (MSI): `{4477F08B-2601-0010-0000-9A09D8342600}`
- FormIt Converter for Revit 2026 (MSI): `{06E56058-9DC2-4B06-8454-D0092F08B9A8}`
- Advance Steel Server Registration for Revit Engine 2026 (MSI):
  `{05BC6921-2026-49D7-A01B-4A9DBE15581D}`
- Autodesk Revit DB Link 2026 (MSI): `{282CD6A9-2601-0010-0000-A6206F572600}`
- OpenStudio CLI For Revit 2026 (MSI): `{EA0B0CD5-0756-43D3-A1D4-D51AAD42D6C2}`
- Revit IFC 2026 (MSI): `{1A9C2C21-2641-4205-0000-992E73C12600}`
- Steel Connections Content for Revit 2026 (MSI):
  `{1DAE3481-2026-46E6-A564-4C485D16FA1D}`
- Autodesk Data Exchange Connector for Revit 2026 (ODIS bundle):
  `{11C6BD4C-0824-3E47-8AD1-DD9B28F18458}`
- US English Content for Revit 2026 (ODIS): `{E07688C3-1A78-3878-AD37-37E99F24BF92}`
- US English Residential Content v2 for Revit 2026 (ODIS):
  `{AAEF7B36-A583-395F-9392-4DC329A3CC23}`
- IFC for Revit 2026 (ODIS): `{EFC3FCBD-7928-338F-8171-AF2621D7C7AE}`
- Revit DB Link for Revit 2026 (ODIS): `{C59850DE-991B-3A18-8ABA-4F1111A0CBEE}`

Orphaned Revit-2026 add-ins/content swept after the core removal (18 total):
Advance Steel Server Registration, Data Exchange Connector, Interoperability
Tools (×2), Issues Addin (×2), Publish NWC Addin (×2), Revit Admin Add-Ins
Manager (×2), Revit DB Link (×2), IFC / Revit IFC (×2), OpenStudio CLI, US
English Content, US English Residential Content v2, Steel Connections Content.

Working ODIS uninstall command (core product):

```
C:\Program Files\Autodesk\AdODIS\V1\installer.exe -i uninstall --trigger_point system ^
  -m C:\ProgramData\Autodesk\ODIS\metadata\{8986CA21-EA9C-32F3-A1DB-C34BD2BDA7A5}\bundleManifest.xml ^
  -x C:\ProgramData\Autodesk\ODIS\metadata\{8986CA21-EA9C-32F3-A1DB-C34BD2BDA7A5}\SetupRes\manifest.xsd ^
  --extension_manifest C:\ProgramData\Autodesk\ODIS\metadata\{8986CA21-EA9C-32F3-A1DB-C34BD2BDA7A5}\setup_ext.xml ^
  --extension_manifest_xsd C:\ProgramData\Autodesk\ODIS\metadata\{8986CA21-EA9C-32F3-A1DB-C34BD2BDA7A5}\SetupRes\manifest_ext.xsd ^
  -o C:\ProgramData\Autodesk\ODIS\metadata\{8986CA21-EA9C-32F3-A1DB-C34BD2BDA7A5}\deploymentCollection.xml
```

The core ODIS teardown took ~28 minutes (it uninstalls each sub-component MSI in
sequence). MSI add-ins take seconds each; ODIS-based add-ins take minutes each.

## Residual folders (Revit-2026-specific, safe to delete after success)

- `%APPDATA%\Autodesk\Revit\Autodesk Revit 2026`
- `%LOCALAPPDATA%\Autodesk\Revit\Autodesk Revit 2026`
- `%APPDATA%\Autodesk\Revit\Addins\2026`
- `%PROGRAMDATA%\Autodesk\Revit\Addins\2026`
- `%PROGRAMDATA%\Autodesk\RVT 2026`
- `%PROGRAMFILES%\Autodesk\Revit 2026`

Runtime guard before any deletion: path must be under an `...\Autodesk\...` tree,
reference `Revit`/`RVT`, and contain `2026`; anything else is refused.

## Commands

```powershell
# Preview only (safe first run) — lists matched products + residual folders:
powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit2026.ps1 -ListOnly

# Full removal (core + orphaned add-ins + residual), unattended and silent:
powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit2026.ps1 -StopRevit -Force

# Core product only:
powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit2026.ps1 -IncludeAddins:$false -RemoveResidualFiles:$false
```

Re-running is idempotent: already-removed items no longer match, and `1605`
("already gone") is treated as success.

## Reinstalling Revit later (reinstall-safe)

This teardown does not block a later reinstall, because it removes products via
Autodesk's own uninstallers (ODIS + msiexec), clearing registrations properly
rather than force-deleting them, and it preserves the installer framework
(ODIS / Autodesk Access), Genuine Service, licensing, and shared libraries a
reinstall depends on. Deleted residual folders are recreated by the installer;
no registry keys are hand-edited (the usual source of reinstall problems).

Steps when reinstalling:

1. **Reboot first** — mandatory only if a run reported exit `3010`, otherwise
   just good hygiene to clear pending file operations.
2. **Install through Autodesk Access / the Autodesk account**, not a stale local
   installer, to get a fresh package and re-add removed content packs.
3. If Autodesk Access still lists the product as installed (UI-cache quirk from
   uninstalling outside Access), refresh/repair or reboot to correct it.

## Reusing for another year

Change `$ProductYear` and `$CorePatterns` in the script. The "Revit + year"
sweep rule, exclusions, resolution order, and residual-path pattern all
parameterize by year with no other changes.