# Uninstall-Revit2026

A scoped, self-elevating PowerShell script that cleanly uninstalls **Autodesk Revit 2026** on Windows — the core application plus its orphaned add-ins, content packs, and exporters — while deliberately preserving shared Autodesk components and other Autodesk products (AutoCAD, Navisworks, other Revit versions).

Autodesk products don't uninstall as a single item. The core application, every add-in, and each content pack register as **separate** entries in Add/Remove Programs, and the core product uses Autodesk's ODIS installer whose command line is unquoted and easy to invoke incorrectly. This script discovers the right entries from the registry, invokes each vendor uninstaller correctly, and stops short of anything shared.

> Built and hardened against a real Revit 2026 install. It is conservative by design: it previews before acting, refuses to touch cross-version or shared components, and logs everything.

## Features

- **Registry-driven discovery** across the 64-bit, 32-bit (WOW6432Node), and per-user uninstall hives — no hardcoded product GUIDs.
- **Correct ODIS invocation.** Runs Autodesk's `AdODIS\V1\installer.exe` directly (not through `cmd`), so its unquoted, space-containing path is handled properly.
- **Multi-method resolution** per product: MSI product code → `QuietUninstallString` → raw `UninstallString`, trying each in order until one succeeds.
- **Precise "Revit + year" sweep** for orphaned add-ins/content, with hard exclusions for shared and cross-version components.
- **Self-elevation** via UAC — launch from a normal shell.
- **Preview mode** (`-ListOnly`) and full `-WhatIf` support.
- **Safe residual cleanup**, gated on a successful uninstall and guarded so it can only ever delete Revit/RVT 2026 folders under an Autodesk tree.
- **Transcript logging** to `%TEMP%`.

## Requirements

- Windows 10/11
- Windows PowerShell 5.1 (built in) — no modules required
- Administrator rights (the script self-elevates via UAC)

## Usage

```powershell
# Preview only — lists matched products and residual folders, changes nothing:
powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit2026.ps1 -ListOnly

# Interactive — prompts before each product and each residual folder:
powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit2026.ps1

# Fully unattended and silent — closes Revit if open, no prompts:
powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit2026.ps1 -StopRevit -Force

# Core application only — skip add-ins and residual cleanup:
powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit2026.ps1 -IncludeAddins:$false -RemoveResidualFiles:$false
```

Run `-ListOnly` first. It is the safety gate: it shows exactly what will be removed before you commit.

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-IncludeAddins` | bool | `$true` | Also remove every product whose name references Revit **and** the target year (add-ins, content, exporters, DB Link, IFC, interop tools). Disable with `-IncludeAddins:$false`. |
| `-RemoveResidualFiles` | bool | `$true` | After a successful uninstall, delete leftover Revit-2026-specific folders (settings, journals, add-in manifests, RVT content, program folder). Disable with `-RemoveResidualFiles:$false`. |
| `-StopRevit` | switch | off | Terminate `Revit.exe` if running. Without it, the script aborts when Revit is open. |
| `-ListOnly` | switch | off | Discover and print matches, then exit. No changes. |
| `-Force` | switch | off | Fully non-interactive: skips per-item prompts **and** suppresses PowerShell's built-in confirmation. |
| `-LogPath` | string | `%TEMP%\...` | Override the transcript log path. |

## How it works

**Product selection.** The core product is matched by name (`Autodesk Revit <year>`). With `-IncludeAddins`, the sweep additionally matches any product whose display name contains both `Revit`/`RVT` and the target year — a single rule that catches the full add-in/content family without an exhaustive list. Cross-version and shared components are excluded: version-range packs (e.g. `2024-2027`), Content Catalog, RealDWG, material libraries, licensing, Genuine Service, Identity Manager, Autodesk Access, ODIS, Desktop Connector, and any version-neutral interop manager. Products without `Revit` in the name (AutoCAD, Navisworks) never match, even when they share the year.

**Uninstall resolution (per product), tried in order:**

1. `msiexec.exe /x {GUID} /qn /norestart` — when the product is Windows Installer-based with a GUID key. Fully silent, deterministic.
2. `QuietUninstallString` — the vendor's own silent command, run directly.
3. Raw `UninstallString` — for EXE uninstallers (Autodesk ODIS), a `--silent` variant is attempted first with the exact vendor command kept as an automatic fallback, so a wrong silent flag can never block the uninstall.

Exit codes `0`, `3010` (reboot required), and `1605` (already gone) are treated as success.

## Safety

- **Preview-first** with `-ListOnly` and full `-WhatIf`.
- **Shared components are never removed** in the default scope.
- **Residual cleanup is gated** on a successful uninstall and constrained by a runtime guard: a path must sit under an `...\Autodesk\...` tree, reference `Revit`/`RVT`, and contain the target year, or it is refused and logged.
- **Clean failure handling:** if a method fails, the script logs the raw uninstall strings and moves on without leaving partial state.

## Logging

Every run writes a full transcript to `%TEMP%\Uninstall-Revit2026_<timestamp>.log`, including each product matched, the exact command invoked, and the exit code. Attach this log when reporting issues.

## Reusing for another Revit year

Change `$ProductYear` and `$CorePatterns` near the top of the script. The sweep rule, exclusions, resolution order, and residual-path pattern all parameterize by year — no other edits required.

## Reinstalling Revit later

This uninstaller is reinstall-safe. It removes products through Autodesk's own
uninstallers (ODIS `installer.exe` and `msiexec` by product code), so product
registrations are cleared properly rather than force-deleted, and it preserves
the Autodesk installer framework (ODIS / Autodesk Access), Genuine Service,
licensing, and shared libraries that a reinstall depends on. Residual folders it
deletes (settings, journals, add-in manifests, content, program folder) are
recreated by the installer.

When you want Revit back:

1. **Reboot first.** Not mandatory unless a run reported exit code `3010`
   (reboot required), but it clears pending file operations and is standard
   pre-reinstall hygiene.
2. **Install through Autodesk Access / your Autodesk account**, not a leftover
   local installer, so you get a fresh package and can re-add the content packs
   that were removed.
3. If Autodesk Access still shows the product as "installed" (a UI-cache quirk
   that can occur when you uninstall outside Access), refresh/repair or reboot
   and it will correct itself.

## Notes and limitations

- The core Revit ODIS uninstall can take a while (it removes each sub-component MSI in sequence); MSI-based add-ins take seconds each.
- Some third-party or ODIS uninstallers display their own progress UI regardless of silent flags. `msiexec` items run fully silent.
- Tested on Windows PowerShell 5.1. PowerShell 7 should work but is not the primary target.
- This tool is not affiliated with or endorsed by Autodesk. "Revit", "AutoCAD", and "Navisworks" are trademarks of Autodesk, Inc.

## Disclaimer

Uninstalling software modifies your system. Review the `-ListOnly` output before running for real, and keep the generated log. The software is provided "as is" — see [LICENSE](LICENSE).

## License

[MIT](LICENSE) © 2026 MrGezz
