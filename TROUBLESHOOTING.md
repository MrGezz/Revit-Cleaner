# Troubleshooting stubborn Autodesk uninstalls

Companion to `Uninstall-Revit.ps1`. When a product refuses to uninstall, the
cause is almost always damaged Windows Installer state on the machine, not the
script. This covers the error codes seen in practice and how to clear them.

## Quick map

| Exit / error | Meaning | What to do |
|---|---|---|
| `1605` | "This action is only valid for a product that is installed" — already gone | Treated as success by the script; nothing to do |
| `3010` | Success, reboot required | Reboot; the removal is complete |
| `1618` | Another install/uninstall is already running | Wait for it (or reboot), then retry |
| `1606` | "Could not access network location …" | Either the MSI's own `DIRCA_INSTALLDIR` composing a relative `INSTALLDIR` (fixed automatically by the `MSI-PropsOverride` attempt) or a broken shell-folder registry value — see below |
| `1603` + Internal Error `2753` | Damaged MSI registration: a custom action sourced from an installed file cannot be resolved | Auto-remediated by the script (neutralize → recache → retry); MS troubleshooter is the manual fallback — see below |

## Error 1603 with Internal Error 2753 (the hard one)

`Internal Error 2753. <file/component key>` means **"the file is not marked for
installation."** A custom action in the product's *uninstall* sequence — one
sourced from an INSTALLED FILE (CustomAction base type 17/18/21/22) — points at
a component whose registration is damaged. This happens after an interrupted
uninstall or a patch that left the registration inconsistent. A plain
`msiexec /x` (product code *or* cached local package path) cannot get past it —
every method returns 1603/2753 — because maintenance mode always runs from the
registered cached package (`Package we're running from ==>` in the verbose
log), so no command-line variation changes the tables being executed.

### Fix A (automated): the script's neutralize → recache → retry chain

`Uninstall-Revit.ps1` (`-NeutralizeBrokenCustomActions`, default on) resolves
this without force-removal, verified end-to-end on the Revit 2023 core
(2026-07-24):

1. Detects `Error 2753` in the failed attempt's verbose log and extracts the
   failing action name.
2. Copies the cached package to
   `%TEMP%\RevitCleanerPatch_<stamp>\<registered name>.msi` — the exact
   registered file name matters: repair source resolution probes
   `SOURCEDIR + <registered PackageName>` and fails 2203/1316 otherwise. A
   pristine backup (`<name>_pristine_<stamp>.msi`) is saved alongside.
3. Sets the broken action's `InstallExecuteSequence` `Condition` to `'0'` in
   the copy (the protected cache itself refuses transacted opens even
   elevated).
4. Recaches the copy with `msiexec /fv` (accepted: PackageCode unchanged;
   carries the INSTALLDIR override because repair costing hits
   `DIRCA_INSTALLDIR` too).
5. Retries by product code — the engine now executes the patched cache, skips
   the dead action, and the rest of the uninstall runs normally with full
   component cleanup and rollback.

### Fix B (manual fallback): Microsoft Program Install and Uninstall Troubleshooter

This tool force-removes the broken registration outright — no component
cleanup, but effective when the automated chain cannot proceed.

1. Download the official package (`.diagcab`):
   `https://download.microsoft.com/download/7/E/9/7E9188C0-2511-4B01-8B4E-0A641EC2F600/MicrosoftProgram_Install_and_Uninstall.meta.diagcab`
   (Reachable via Microsoft's support topic "Fix problems that block programs
   from being installed or removed".)
2. Run it → **Next** → choose **Uninstalling**.
3. Pick the stuck product (e.g. **Revit 2023**) from the list. If it is not
   listed, choose it by **product code** — for the 2023 core that is
   `{7346B4A0-2300-0510-0000-705C0D862004}`.
4. Let it remove the registration and clean the broken cache entry.
5. **Re-run the script** for that year, e.g.
   `powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit.ps1 -ProductYear 2023 -StopRevit -Force`.
   The core now reports `1605` ("already gone") → success, so the run completes
   and residual cleanup finally proceeds.

Windows 11 note: Microsoft positions this troubleshooter for Windows 10, but the
`.diagcab` still runs on Windows 11. The built-in **Settings → Apps → Installed
apps → (…) → Uninstall/Repair** is the sanctioned Win11 path but will hit the
same 2753 on a damaged package; the troubleshooter is the one that force-clears
it.

### Diagnose first (optional)

Capture a verbose log to see exactly which file 2753 references:

```
msiexec /x {7346B4A0-2300-0510-0000-705C0D862004} /qn /norestart /l*v "%TEMP%\revit2023_v.log"
```

Open the log and search for `2753` and `Return value 3`. If the referenced file
still exists and the original source media is available, a repair-then-uninstall
(`msiexec /fvomus {code}` then `/x {code}`) can re-cache the package and let the
uninstall succeed. For Autodesk, source media is usually gone, so the
troubleshooter route is faster.

Two cautions learned on ICECREAMASSASIN:

- Forcing `INSTALLDIR`/`ROOTDRIVE` on an uninstall can itself provoke 2753 by
  flipping a component condition out of the action sequence — this is why the
  script keeps `MSI-PropsOverride` as the LAST MSI attempt, after the
  unmodified ones.
- Machine state drifts. The Revit 2023 core returned 1603/2753 in the morning
  runs and pure 1606-at-CostFinalize by the final runs of the same day. Always
  diagnose from the *newest* per-attempt log
  (`%TEMP%\MSIVerbose_<guid>_<stamp>_<Kind>.log`) — the script writes one per
  attempt precisely so an earlier attempt's evidence is never overwritten.

## Error 1606 ("Could not access network location …")

Two distinct causes produce this message. Identify yours from the verbose log:
if the "network location" is a bare **relative fragment** like `Revit 2023\`
and the failure lands in `CostFinalize` (with `Note: 1: 1314`), it is Cause A;
if it names a real (dead) drive/UNC path, it is Cause B.

### Cause A — the MSI's own `DIRCA_INSTALLDIR` action (proven, Revit 2023 core)

Autodesk's Revit MSIs carry a Type-51 custom action:

```
DIRCA_INSTALLDIR:  INSTALLDIR = [INSTALLDIR][ADSK_INSTALL_PATH]\
                   condition: NOT INSTALLDIR><ADSK_INSTALL_PATH   (>< = "contains")
ADSK_INSTALL_PATH: "Revit <year>"   (Property table default)
```

At install time ODIS passes the parent folder in `INSTALLDIR`, so the
composition yields a real absolute path. Uninstalling **directly with msiexec**
(the registry `UninstallString` is plain `MsiExec.exe /X{code}`), `INSTALLDIR`
arrives empty, the action composes the bare relative fragment `Revit <year>\`,
and `CostFinalize` fails: `Note: 1: 1314` → `Error 1606. Could not access
network location Revit <year>\.` — surfaced to the caller only as generic exit
1603. This is **not** a SourceList problem (the SourceList was verified healthy
while this fired), and no amount of LocalPackage/SourceList surgery fixes it.

**Fix:** pass an absolute `INSTALLDIR` that *contains* the `Revit <year>`
token. The contains-condition then evaluates true, `DIRCA_INSTALLDIR` is
skipped, and the override survives to costing:

```
msiexec /x {code} /qn /norestart ROOTDRIVE=C:\ INSTALLDIR="C:\Program Files\Autodesk\Revit <year>"
```

The script's `MSI-PropsOverride` attempt issues exactly this automatically
(last among the MSI attempts — see the 2753 cautions above). Formatting
matters: no trailing backslash before the closing quote (`\"` escapes the
quote and mangles the rest of the command line), and `ROOTDRIVE` unquoted
(no spaces; MSI requires its trailing backslash).

### Cause B — broken shell-folder registry value

A **shell-folder registry value** points at an invalid, redirected, or blank
path, and the installer can't resolve it. It recurs across packages until the
underlying value is fixed.

Permanent fix — correct the offending value (back up the key first):

- `HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders`
- `HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders`
- `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders`

Look for an entry (commonly `Personal`, `AppData`, `Common AppData`, `Cache`, or
a `{GUID}`) whose data points to a drive/UNC path that no longer exists or is
empty, and restore it to the correct local path (e.g. `Personal` →
`%USERPROFILE%\Documents`). Sign out/in afterward.

## General order of operations

1. Run `-ListOnly` to confirm scope.
2. Run the uninstall (`-StopRevit -Force`).
3. If a product fails, read the transcript `%TEMP%\Uninstall-Revit<year>_*.log`
   (the script prints the raw uninstall strings for any failure) and the
   per-attempt verbose logs `%TEMP%\MSIVerbose_<guid>_<stamp>_<Kind>.log`.
4. `1606` → Cause A is handled automatically by the `MSI-PropsOverride`
   attempt; for Cause B, fix the shell-folder values for a permanent cure.
5. `1603` / `2753` → Microsoft Program Install and Uninstall Troubleshooter, then
   re-run the script so residual cleanup completes.
