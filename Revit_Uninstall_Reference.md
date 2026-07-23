# Autodesk Revit Uninstall — Teardown Reference

Working knowledge for scripting a clean, scoped uninstall of an Autodesk Revit
product on Windows without disturbing shared components or other Autodesk apps.
Derived from a verified Revit 2026 removal on machine `ICECREAMASSASIN`
(2026-07-17), then generalized to any year. The delivered script is
`Uninstall-Revit.ps1` (PowerShell 5.1 compatible, self-elevating,
year-parameterized via `-ProductYear`, default 2026). Published to GitHub with
an MIT LICENSE and a README (author: MrGezz). Current revision: 2026-07-24
(MSI-LocalPackage + MSI-PropsOverride resolution, 1606 root-cause fix, rebuilt
self-elevation, automated 2753 neutralize→recache→retry — see lessons 7–12 and
the troubleshooting section).

## What worked

The core Revit 2026 product uninstalled with exit 0 via Autodesk's ODIS
installer. The MSI-based add-ins uninstalled instantly by product code.
Shared components (material libraries, licensing, RealDWG, Content Catalog)
were preserved, and AutoCAD/Navisworks 2026 (separate apps) were untouched.

## Twelve lessons that cost debugging cycles

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

6. **Self-elevation must pass a single command-line STRING to `Start-Process`,
   not an array, when the script path can contain spaces.** In Windows
   PowerShell 5.1, `Start-Process -ArgumentList @(...)` re-quotes array elements
   and mangles a `-File` path like `E:\ICZ 2\Desktop\Uninstall-Revit2026.ps1`,
   so the elevated relaunch silently fails to find the script (self-elevation
   "does nothing" from a non-admin shell, while an already-admin shell works
   because the relaunch path is skipped). Fix: build the whole command line as
   one pre-quoted string —
   `-NoProfile -ExecutionPolicy Bypass -File "<path>" <args>` — and pass that to
   `-ArgumentList`. Also guard an empty `$PSCommandPath` (dot-sourced/pasted),
   treat a thrown `Start-Process` as a cancelled UAC prompt, and default the
   child `ExitCode` to 0 (it is null for ShellExecute/RunAs launches). Note the
   elevated relaunch opens a *separate* window that closes on completion — read
   the `%TEMP%` log for results, not the original window.

7. **`powershell.exe -File` cannot bind `[bool]` parameters at all (nor
   `-Switch:$false`).** PS 5.1 passes every `-File` argument as a literal string
   and rejects it with "Boolean parameters accept only Boolean values and
   numbers" — verified on this machine for `True`, `False`, `1`, `0`, and
   `$false`. The elevated child died at parameter binding *before*
   `Start-Transcript`: no log, elevation "does nothing". Amends lesson 6: the
   single pre-quoted string is still right, but it must be a `-Command` line —
   `-Command "& '<single-quoted path>' <args>; exit $LASTEXITCODE"`. Bools then
   parse natively and a spaced path survives. The trailing
   `; exit $LASTEXITCODE` is mandatory: in `-Command` mode an `exit N` inside
   the invoked *script* only sets `$LASTEXITCODE`, and the child collapses every
   non-zero script exit to 1 (verified: `exit 42` → 1 without it, 42 with it).

8. **Do not GUID-collapse msiexec candidates that carry `PROPERTY=` overrides.**
   De-duping msiexec attempts by extracted GUID alone silently deleted the
   `MSI-PropsOverride` candidate (same GUID as the plain `/x {guid}` attempt),
   so the 1606 workaround never ran — transcripts show only two methods ever
   executing. Candidates with property assignments are functionally different
   commands; key them by their full normalized command line.

9. **Never end a quoted msiexec property value with a backslash.**
   `INSTALLDIR="C:\...\Revit 2023\"` — the `\"` escapes the closing quote and
   mangles every argument after it. Drop the trailing backslash inside quotes
   (MSI normalizes it back), and pass `ROOTDRIVE=C:\` unquoted (no spaces, and
   MSI requires *its* trailing backslash).

10. **One `/L*V` log per attempt.** `/L*V` truncates on open, so attempts
    sharing a filename destroy each other's evidence — the LocalPackage and
    bare-MSI attempts of one run were observed writing the identical file.
    Name logs `MSIVerbose_<guid>_<stamp>_<Kind>.log`.

11. **Maintenance mode always runs from the REGISTERED cached package.**
    `msiexec /x <path>.msi` uses the path only to identify the product — the
    verbose log's `Package we're running from ==>` line shows the registered
    cache, so table edits in a side copy are never seen. The working route for
    delivering an edited database (proven on the Revit 2023 core,
    2026-07-24): edit a %TEMP% copy → `msiexec /fv "<copy>"` to RECACHE it
    (accepted because the PackageCode is unchanged; carry the INSTALLDIR
    override since repair costing hits DIRCA_INSTALLDIR too) → then a plain
    product-code `/x`. Two hard requirements: the copy MUST keep the
    registered package's exact FILE NAME (repair source resolution probes
    `SOURCEDIR + <registered PackageName>` and fails 2203/1316 + SECREPAIR
    otherwise — use a per-run subfolder), and the registered language
    transform must not touch the edited row (verify via `ApplyTransform`
    before betting on it). Also: C:\Windows\Installer cannot be edited in
    place — transacted opens are refused there even elevated.

12. **PowerShell 5.1 + Windows Installer COM has three traps, all observed
    live.** (a) `InvokeMember('OpenDatabase', ...)` can throw
    DISP_E_TYPEMISMATCH where DIRECT dispatch (`$installer.OpenDatabase()`)
    succeeds on identical arguments — try direct first, keep InvokeMember as
    fallback. (b) Direct-dispatch VOID COM calls (`Execute`, `Close`,
    `Commit`) emit `$null` into the function's output stream, so an
    uncaptured sequence of them turns the function's return value into an
    array (symptom: `/x "    C:\...msi"` with leading spaces → relative-path
    resolution → 1619). `[void]`-cast every one. (c) RCWs keep the database's
    FILE HANDLE open until finalized — ReleaseComObject on every wrapper
    (including fetched Records and replaced Views) plus
    `[GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()`
    before handing the file to msiexec, or it fails 1619 with
    0x80030020 STG_E_SHAREVIOLATION.

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

1. MSI-LocalPackage — `msiexec.exe /x "C:\Windows\Installer\<cached>.msi"` —
   uninstall from the cached local package (resolved via the Windows Installer
   COM `LocalPackage` property), bypassing SourceList/network-source resolution.
2. MSI product code — `msiexec.exe /x {GUID} /qn /norestart` — when
   `WindowsInstaller = 1` and the registry key name is a GUID. Most deterministic.
3. MSI-PropsOverride — product code plus `ROOTDRIVE=C:\ INSTALLDIR="<ProgramFiles>\
   Autodesk\Revit <year>"` — deliberately LAST among the MSI attempts because
   forcing directory properties on an uninstall can flip component conditions
   (2753 risk), but it is the attempt that clears the DIRCA_INSTALLDIR 1606
   (see troubleshooting below).
4. `QuietUninstallString` (run the exe directly, not via cmd).
5. Raw `UninstallString` (run the exe directly; coerce msiexec `/I`→`/X` and add
   `/qn /norestart`; for ODIS EXE uninstallers, try `--silent` first with the
   exact vendor command as fallback). This is the ODIS path for the core product.

Every MSI attempt writes its own verbose log:
`%TEMP%\MSIVerbose_<guid>_<stamp>_<Kind>.log`.

When an MSI attempt exits 1603 and its log shows **Internal Error 2753**, the
script auto-remediates (`-NeutralizeBrokenCustomActions`, default on): it
copies the cached package to `%TEMP%\RevitCleanerPatch_<stamp>\<registered
name>.msi` (pristine backup saved alongside as `<name>_pristine_<stamp>.msi`),
conditions the named action out (`'0'`), recaches via `/fv`, and retries by
product code (`MSI-Recached`). Bounded at 5 repairs per method. This is what
finally removed the Revit 2023 core on 2026-07-24 (exit 0 + full residual
cleanup; a re-run confirmed the product gone).

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
powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit.ps1 -ProductYear 2026 -ListOnly

# Full removal (core + orphaned add-ins + residual), unattended and silent:
powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit.ps1 -ProductYear 2026 -StopRevit -Force

# Core product only:
powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit.ps1 -ProductYear 2026 -IncludeAddins:$false -RemoveResidualFiles:$false
```

Re-running is idempotent: already-removed items no longer match, and `1605`
("already gone") is treated as success. From a non-admin shell the script
self-elevates (one UAC prompt) and runs in a separate window; watch the log.

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

## Troubleshooting stubborn products (see TROUBLESHOOTING.md)

Failures are almost always damaged Windows Installer state on the machine, not a
script bug. Codes seen in practice:

- `1605` already gone (success), `3010` reboot-required (success), `1618`
  another install running (retry).
- **`1606`** "Could not access network location <x>\" — two distinct causes:
  1. **(proven on this machine, Revit 2023 core)** The MSI's own Type-51 action
     `DIRCA_INSTALLDIR` composes `INSTALLDIR = [INSTALLDIR][ADSK_INSTALL_PATH]\`
     (condition `NOT INSTALLDIR><ADSK_INSTALL_PATH`, `><` = "contains";
     `ADSK_INSTALL_PATH = "Revit <year>"`). Uninstalling directly with msiexec
     (outside ODIS) INSTALLDIR arrives empty, so it becomes the bare relative
     fragment `Revit <year>\` and CostFinalize dies (Note 1314 →
     `Error 1606. Could not access network location Revit <year>\`), surfaced
     only as generic exit 1603. **Fix:** pass an absolute `INSTALLDIR`
     *containing* the `Revit <year>` token — the contains-condition goes false,
     the CA skips, the override survives. The script's `MSI-PropsOverride`
     attempt does exactly this. It is **not** a SourceList problem — the
     SourceList was verified healthy while this fired.
  2. A broken shell-folder registry value (`User Shell Folders` /
     `Shell Folders`) pointing at an invalid or blank path. Permanent cure is
     fixing the offending registry value (see TROUBLESHOOTING.md).
- **`1603` + Internal Error `2753`** ("the file is not marked for installation")
  — a custom action sourced from an installed file whose component registration
  is damaged (interrupted uninstall or bad patch). A plain `msiexec /x`
  (product code or cached-package path) cannot get past it because maintenance
  mode always executes the REGISTERED cache (lesson 11). **Fix (automated):**
  the script's neutralize → `/fv` recache → product-code retry chain
  (`-NeutralizeBrokenCustomActions`, default on) — this is what removed the
  Revit 2023 core on 2026-07-24. **Fallback (manual):** Microsoft Program
  Install and Uninstall Troubleshooter
  (`MicrosoftProgram_Install_and_Uninstall.meta.diagcab`) → Uninstalling →
  force-remove, then re-run the script (core returns 1605 = success). Note:
  forcing `INSTALLDIR`/`ROOTDRIVE` on an uninstall can itself provoke 2753 by
  flipping component conditions — which is why `MSI-PropsOverride` runs LAST.
  Machine state also drifts: a product that returned 2753 one day can return
  plain 1606 later; always diagnose from the *newest* `MSIVerbose_*_<Kind>.log`.

Verified case: `-ProductYear 2023` on ICECREAMASSASIN (2026-07-23) removed all 12
Navisworks 2023 exporters cleanly but `Revit 2023` core kept failing 1603
(earlier in the day as 2753; by the final runs as pure 1606 at CostFinalize —
machine state drifted). Root cause decoded from the cached MSI
(`C:\Windows\Installer\1d05a2.msi`): the `DIRCA_INSTALLDIR` mechanism above.
The dedup defect (lesson 8) had silently removed the `MSI-PropsOverride`
attempt, so the fix never ran. Revit 2023 core product code:
`{7346B4A0-2300-0510-0000-705C0D862004}`. **Resolved 2026-07-24:** after the
lesson 11/12 fixes, the neutralize → `/fv` recache → product-code retry chain
removed the core with exit 0, residual cleanup ran, and a verification re-run
found no Revit 2023 products remaining. The Microsoft troubleshooter was never
needed.

Folded into the repo copy 2026-07-23 (supersedes canonical v7): the
`MSI-LocalPackage` method, the `MSI-PropsOverride` last-resort, SourceList
dump/purge helpers, plus seven fixes — dedup no longer GUID-collapses
property-carrying candidates (lesson 8); trailing-backslash quoting corrected
(lesson 9); self-elevation rebuilt on `-Command` with exit-code propagation
(lesson 7); per-attempt verbose MSI logs (lesson 10); registry surgery
(`Repair-MsiUserDataCache` + SourceList purge) moved after the Read-Host /
ShouldProcess consent gates; StrictMode-safe `$installer = $null` pre-init
before the COM try/finally blocks; `Repair-MsiUserDataCache` reuses
`Get-MsiSquishedGuid` instead of duplicating the squish algorithm.

## Repo

Lives at `C:\Users\IceCreamAssasin\source\repos\Revit-Cleaner` (git). Files:
`Uninstall-Revit.ps1`, `README.md`, `LICENSE` (MIT), `TROUBLESHOOTING.md`, and
this reference (`Revit_Uninstall_Reference.md`). User handles git commits manually. The old
`Uninstall-Revit2026.ps1` was renamed to `Uninstall-Revit.ps1`
(`git rm`/`git mv` the stale name).

## Targeting another year

Pass `-ProductYear <yyyy>` (e.g. `-ProductYear 2024`); default is `2026`. The
core match, the "Revit + year" add-in sweep, the residual folders, the residual
path guard, and the self-elevation relaunch are all scoped to that value — no
edits required. The exclusions (version-range packs, RealDWG, Content Catalog,
shared/version-neutral components) are year-agnostic and always apply.
