<#
.SYNOPSIS
    Uninstalls the Autodesk Revit 2026 core product (Revit only) on Windows.

.DESCRIPTION
    Discovers the installed Autodesk Revit 2026 product by reading the Windows
    "Uninstall" registry hives (64-bit, 32-bit/WOW6432Node, and per-user), then
    runs the vendor-registered uninstaller silently.

    Scope is deliberately conservative: by default ONLY the core Revit 2026
    application is removed. Shared Autodesk components (Material Libraries,
    Content/Template Libraries, Licensing / Desktop Licensing Service, Genuine
    Service, Identity Manager, ODIS installer, Autodesk Access) are left in
    place because other Autodesk products (Civil 3D, AutoCAD, Navisworks, etc.)
    depend on them. Revit-specific 2026 add-ins (Batch Print, eTransmit, Model
    Review, Coordination Model, FormIt Converter, Personal Accelerator, Cloud
    Models for Revit) are only removed when -IncludeAddins is supplied.

    Resolution order for each product's uninstall command:
        1. QuietUninstallString (vendor-provided silent command)
        2. msiexec /x {ProductCode} /qn /norestart   (when WindowsInstaller = 1)
        3. UninstallString (interactive fallback; /qn appended for msiexec)

.PARAMETER IncludeAddins
    Also remove every product whose name references Revit and the target year
    (2026) - add-ins, content packs, exporters, DB Link, IFC, interop tools,
    etc. - which are orphaned once the core application is gone. Default: $true.
    Disable with -IncludeAddins:$false. Cross-version and shared components
    (Content Catalog 2024-2027, RealDWG, version-neutral interop managers,
    material libraries, licensing) are always preserved.

.PARAMETER RemoveResidualFiles
    After uninstall, delete leftover Revit 2026-specific folders (per-user
    settings/journals, add-in manifests, RVT 2026 content/templates, and any
    residual Revit 2026 program folder). Default: $true. Disable with
    -RemoveResidualFiles:$false. A runtime guard only permits deletion of
    paths under an Autodesk tree that reference Revit/RVT and 2026; shared
    Autodesk trees are never removed.

.PARAMETER StopRevit
    If Revit.exe is running, terminate it before uninstalling. Without this
    switch the script aborts when Revit is running (safer default).

.PARAMETER ListOnly
    Discover and print matching products, then exit. Performs no changes.
    Equivalent to a dry run.

.PARAMETER Force
    Fully non-interactive: skips the per-product Read-Host prompt AND suppresses
    PowerShell's built-in ShouldProcess "Are you sure?" confirmation (this script
    is ConfirmImpact=High, which would otherwise prompt even under -Force).

.PARAMETER LogPath
    Full path for the transcript log. Defaults to
    %TEMP%\Uninstall-Revit2026_<timestamp>.log

.EXAMPLE
    # Preview everything that would be removed (products + residual folders), no changes:
    powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit2026.ps1 -ListOnly

.EXAMPLE
    # Uninstall Revit 2026 + Revit-only add-ins + residual files, prompting each step:
    powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit2026.ps1

.EXAMPLE
    # Fully unattended, close Revit if open, no prompts:
    powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit2026.ps1 -StopRevit -Force

.EXAMPLE
    # Core product only (skip add-ins and residual cleanup):
    powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit2026.ps1 -IncludeAddins:$false -RemoveResidualFiles:$false

.NOTES
    Requires an elevated (Administrator) session; the script self-elevates.
    Exit code 0 = success, 3010 = success (reboot required), 3 = partial failure,
    2 = nothing found, 1 = aborted.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [bool]$IncludeAddins       = $true,
    [bool]$RemoveResidualFiles = $true,
    [switch]$StopRevit,
    [switch]$ListOnly,
    [switch]$Force,
    [string]$LogPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# -Force implies fully non-interactive: suppress PowerShell's own ShouldProcess
# confirmation (ConfirmImpact=High would otherwise prompt "Are you sure?" for
# every product even under -Force).
if ($Force) { $ConfirmPreference = 'None' }

# --- Configuration --------------------------------------------------------
$ProductYear = '2026'

# Core product: always targeted.
$CorePatterns = @(
    'Autodesk Revit 2026'
)

# Revit-only add-ins: targeted only with -IncludeAddins.
$AddinPatterns = @(
    'Batch Print*Revit 2026',
    'eTransmit*Revit 2026',
    'Model Review*2026',
    'Coordination Model*Revit 2026',
    'FormIt Converter*Revit 2026',
    'Personal Accelerator for Revit*2026',
    'Cloud Models for Revit 2026',
    'Autodesk Revit Content Libraries 2026',   # Revit-specific content, not the shared material lib
    'Revit*2026*Add-in'
)

# Shared / cross-product components: NEVER touched in this mode.
$SharedExclusions = @(
    '*Material Library*',
    '*Advanced Material Library*',
    '*Shared Components*',
    '*Licensing*',
    '*License*',
    '*Desktop Licensing Service*',
    '*Genuine Service*',
    '*Identity Manager*',
    '*Autodesk Access*',
    '*Desktop App*',
    '*Autodesk Access Core*',
    '*Single Sign On*',
    '*ODIS*',
    '*AdODIS*',
    '*Autodesk Installer*',
    '*Content Catalog*',
    '*RealDWG*',
    '*Interoperability Engine Manager*',
    '*Desktop Connector*'
)

# Revit 2026-specific residual folders removed only with -RemoveResidualFiles.
# Every entry references a Revit/RVT 2026 path; shared Autodesk trees are never
# listed here, and a runtime guard re-verifies each path before deletion.
$ResidualPaths = @(
    (Join-Path $env:APPDATA        'Autodesk\Revit\Autodesk Revit 2026'),
    (Join-Path $env:LOCALAPPDATA   'Autodesk\Revit\Autodesk Revit 2026'),
    (Join-Path $env:APPDATA        'Autodesk\Revit\Addins\2026'),
    (Join-Path $env:ProgramData    'Autodesk\Revit\Addins\2026'),
    (Join-Path $env:ProgramData    'Autodesk\RVT 2026'),
    (Join-Path ${env:ProgramFiles} 'Autodesk\Revit 2026')
)

# --- Self-elevation -------------------------------------------------------
function Test-IsAdministrator {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdministrator)) {
    Write-Host 'Elevation required. Relaunching as Administrator...' -ForegroundColor Yellow

    # Guard: self-elevation needs the script's own path. It is empty when the
    # script is dot-sourced or pasted into the console rather than run as a file.
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        Write-Host 'Cannot self-elevate: script path is unknown. Run it with -File (not dot-sourced/pasted), or start an elevated PowerShell first.' -ForegroundColor Red
        exit 1
    }

    # Build the relaunch command line as a SINGLE string, not an array.
    # Start-Process re-quotes array elements in Windows PowerShell 5.1 and
    # mangles a script path that contains spaces (e.g. "E:\ICZ 2\Desktop\..."),
    # which silently breaks self-elevation. A single pre-quoted string is passed
    # to the child verbatim.
    $passArgs = @()
    $passArgs += ('-IncludeAddins:{0}'       -f $IncludeAddins)
    $passArgs += ('-RemoveResidualFiles:{0}' -f $RemoveResidualFiles)
    if ($StopRevit) { $passArgs += '-StopRevit' }
    if ($ListOnly)  { $passArgs += '-ListOnly' }
    if ($Force)     { $passArgs += '-Force' }
    if ($LogPath)   { $passArgs += ('-LogPath "{0}"' -f $LogPath) }

    $cmdLine = '-NoProfile -ExecutionPolicy Bypass -File "{0}" {1}' -f $PSCommandPath, ($passArgs -join ' ')

    try {
        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $cmdLine -Verb RunAs -PassThru -Wait -ErrorAction Stop
    }
    catch {
        # Thrown when the user clicks No / cancels the UAC prompt, or UAC is blocked.
        Write-Host "Elevation was cancelled or blocked at the UAC prompt: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }

    # ExitCode can be null for ShellExecute (RunAs) launches; default to 0.
    $code = 0
    if ($proc -and $null -ne $proc.ExitCode) { $code = $proc.ExitCode }
    exit $code
}

# --- Logging --------------------------------------------------------------
if (-not $LogPath) {
    $stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $LogPath = Join-Path $env:TEMP "Uninstall-Revit2026_$stamp.log"
}
try { Start-Transcript -Path $LogPath -Append | Out-Null } catch { }

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'HH:mm:ss'
    $color = switch ($Level) {
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        'OK'    { 'Green' }
        default { 'Gray' }
    }
    Write-Host ("[{0}] {1,-5} {2}" -f $ts, $Level, $Message) -ForegroundColor $color
}

Write-Log "Revit $ProductYear uninstall started. Log: $LogPath"
Write-Log ("Mode: {0}{1}{2}{3}" -f `
    $(if ($ListOnly) { 'ListOnly ' } else { 'Uninstall ' }),
    $(if ($IncludeAddins) { '+Add-ins ' } else { 'CoreOnly ' }),
    $(if ($RemoveResidualFiles) { '+Residual ' } else { '' }),
    $(if ($Force) { 'Force' } else { 'Interactive' }))

# --- Product discovery ----------------------------------------------------
# StrictMode-safe property reader: returns the value or $null, never throws
# when the registry key lacks the requested value.
function Get-Prop {
    param($Obj, [string]$Name)
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -ne $p) { return $p.Value }
    return $null
}

function Get-InstalledPrograms {
    $hives = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    foreach ($hive in $hives) {
        if (-not (Test-Path $hive)) { continue }
        Get-ChildItem -Path $hive -ErrorAction SilentlyContinue | ForEach-Object {
            $props = $null
            try { $props = Get-ItemProperty -Path $_.PsPath -ErrorAction Stop } catch { return }
            $dn = Get-Prop $props 'DisplayName'
            if ([string]::IsNullOrWhiteSpace($dn)) { return }
            [pscustomobject]@{
                DisplayName          = $dn
                DisplayVersion       = Get-Prop $props 'DisplayVersion'
                Publisher            = Get-Prop $props 'Publisher'
                UninstallString      = Get-Prop $props 'UninstallString'
                QuietUninstallString = Get-Prop $props 'QuietUninstallString'
                WindowsInstaller     = [int](Get-Prop $props 'WindowsInstaller')
                KeyName              = $_.PSChildName
                RegistryPath         = $_.PsPath
            }
        }
    }
}

function Test-MatchesAny {
    param([string]$Value, [string[]]$Patterns)
    foreach ($p in $Patterns) { if ($Value -like $p) { return $true } }
    return $false
}

$all = @(Get-InstalledPrograms | Where-Object { $_.Publisher -like '*Autodesk*' -or $_.DisplayName -like '*Revit*' })

# Core product is always targeted. With -IncludeAddins, ALSO sweep every
# product whose name references Revit AND the target year - the add-ins,
# content packs, exporters, DB Link, IFC, interop tools, etc. that install as
# separate products and are orphaned once the core application is gone.
# Cross-version and shared components (Content Catalog 2024-2027, RealDWG,
# version-neutral interop managers with no year) are excluded so other Revit
# versions keep working.
$targets = @(
    $all |
    Where-Object {
        $name = $_.DisplayName
        $isCore = Test-MatchesAny -Value $name -Patterns $CorePatterns
        $isRevitYear = $IncludeAddins -and
                       ($name -match '(?i)\bRevit\b|\bRVT\b') -and
                       ($name -like "*$ProductYear*")
        $isCore -or $isRevitYear
    } |
    Where-Object { -not (Test-MatchesAny -Value $_.DisplayName -Patterns $SharedExclusions) } |
    Where-Object { $_.DisplayName -notmatch '\d{4}\s*[-–]\s*\d{4}' } |
    Sort-Object DisplayName -Unique
)

if ($targets.Count -eq 0) {
    Write-Log "No Autodesk Revit $ProductYear products found in the uninstall registry." 'WARN'
    if ($all.Count -gt 0) {
        Write-Log "Autodesk/Revit entries present on this machine (for reference):"
        $all | Sort-Object DisplayName -Unique | ForEach-Object { Write-Log "    - $($_.DisplayName)" }
    }
    try { Stop-Transcript | Out-Null } catch { }
    exit 2
}

Write-Log "Matched $($targets.Count) product(s) for removal:" 'OK'
$targets | ForEach-Object { Write-Log ("    - {0}  [{1}]" -f $_.DisplayName, $_.DisplayVersion) }

if ($ListOnly) {
    if ($RemoveResidualFiles) {
        $existing = @($ResidualPaths | Where-Object { $_ -and (Test-Path -LiteralPath $_) })
        if ($existing.Count -gt 0) {
            Write-Log 'Residual folders that would be removed:'
            $existing | ForEach-Object { Write-Log "    - $_" }
        }
        else {
            Write-Log 'No Revit 2026 residual folders found.'
        }
    }
    Write-Log 'ListOnly specified - no changes made.' 'OK'
    try { Stop-Transcript | Out-Null } catch { }
    exit 0
}

# --- Running-process guard ------------------------------------------------
$revitProcs = @(Get-Process -Name 'Revit' -ErrorAction SilentlyContinue)
if ($revitProcs.Count -gt 0) {
    if ($StopRevit) {
        Write-Log "Revit.exe is running - terminating $($revitProcs.Count) process(es)." 'WARN'
        $revitProcs | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    }
    else {
        Write-Log 'Revit.exe is running. Close Revit or re-run with -StopRevit. Aborting.' 'ERROR'
        try { Stop-Transcript | Out-Null } catch { }
        exit 1
    }
}

# --- Uninstall command resolution -----------------------------------------
# Split a registry command line into executable + argument string WITHOUT
# routing through cmd.exe. cmd mangles an unquoted, space-containing path such
# as  C:\Program Files\Autodesk\AdODIS\V1\installer.exe  (it reads the exe as
# "C:\Program"), which is exactly what produced the generic "exit 1" on
# Autodesk's ODIS uninstaller. Start-Process quotes -FilePath correctly.
function Split-Command {
    param([string]$CommandLine)
    $cl = $CommandLine.Trim()

    # Quoted executable: take whatever is between the first pair of quotes.
    if ($cl.StartsWith('"')) {
        $end = $cl.IndexOf('"', 1)
        if ($end -lt 1) { return [pscustomobject]@{ File = $cl.Trim('"'); Args = '' } }
        return [pscustomobject]@{
            File = $cl.Substring(1, $end - 1)
            Args = $cl.Substring($end + 1).Trim()
        }
    }

    # Unquoted: split on the first '.exe' token so spaced paths survive intact.
    $m = [regex]::Match($cl, '(?i)^(.*?\.exe)(?:\s+(.*))?$')
    if ($m.Success) {
        return [pscustomobject]@{
            File = $m.Groups[1].Value.Trim()
            Args = $m.Groups[2].Value.Trim()
        }
    }

    # Last resort: split on the first space.
    $sp = $cl.IndexOf(' ')
    if ($sp -lt 0) { return [pscustomobject]@{ File = $cl; Args = '' } }
    return [pscustomobject]@{ File = $cl.Substring(0, $sp); Args = $cl.Substring($sp + 1).Trim() }
}

# Build an ordered list of uninstall attempts. If one method fails the caller
# falls through to the next.
function Get-UninstallCandidates {
    param($Product)
    $list = @()

    # 1) MSI product code (deterministic) when the registry key is a GUID.
    if ($Product.WindowsInstaller -eq 1 -and $Product.KeyName -match '^\{[0-9A-Fa-f\-]{36}\}$') {
        $list += [pscustomobject]@{ File = 'msiexec.exe'; Args = "/x $($Product.KeyName) /qn /norestart"; Kind = 'MSI' }
    }

    # 2) Vendor-provided silent command.
    if (-not [string]::IsNullOrWhiteSpace($Product.QuietUninstallString)) {
        $s = Split-Command $Product.QuietUninstallString
        $list += [pscustomobject]@{ File = $s.File; Args = $s.Args; Kind = 'Quiet' }
    }

    # 3) Raw UninstallString. For msiexec, coerce /I->/X and add silent flags.
    #    For an EXE uninstaller (e.g. Autodesk ODIS installer.exe), try a silent
    #    variant first, then the exact vendor string as a guaranteed fallback so
    #    a wrong silent flag can never block the uninstall.
    if (-not [string]::IsNullOrWhiteSpace($Product.UninstallString)) {
        $raw = $Product.UninstallString
        if ($raw -match '(?i)msiexec') {
            $raw = $raw -replace '(?i)/I(\{[0-9A-Fa-f\-]{36}\})', '/X$1'
            if ($raw -notmatch '(?i)/qn|/quiet') { $raw = "$raw /qn" }
            if ($raw -notmatch '(?i)/norestart') { $raw = "$raw /norestart" }
            $s = Split-Command $raw
            $list += [pscustomobject]@{ File = $s.File; Args = $s.Args; Kind = 'Fallback' }
        }
        else {
            $s = Split-Command $raw
            if ($s.Args -notmatch '(?i)(^|\s)(--silent|-silent|/silent|-q|/q)(\s|$)') {
                $list += [pscustomobject]@{ File = $s.File; Args = ($s.Args + ' --silent').Trim(); Kind = 'Silent' }
            }
            $list += [pscustomobject]@{ File = $s.File; Args = $s.Args; Kind = 'Fallback' }
        }
    }

    # De-duplicate, collapsing msiexec variants that target the same GUID
    # (e.g. "msiexec.exe /x {g}" and "MsiExec.exe /X{g}") into a single attempt.
    $seen = @{}
    $unique = @()
    foreach ($c in $list) {
        $norm = ('{0} {1}' -f $c.File, $c.Args).ToLowerInvariant()
        $guid = [regex]::Match($norm, '\{[0-9a-f\-]{36}\}')
        if ($norm -match 'msiexec' -and $guid.Success) { $key = 'msi:' + $guid.Value }
        else { $key = $norm }
        if (-not $seen.ContainsKey($key)) { $seen[$key] = $true; $unique += $c }
    }
    # NOTE: plain 'return $unique' (NOT ',$unique'). The unary-comma wrapper
    # double-nests the array so the caller's foreach iterates once over the
    # whole array, collapsing all candidates into one object with array-valued
    # properties -> Start-Process gets an array for -FilePath and throws.
    return $unique
}

# --- Execution ------------------------------------------------------------
$successCodes = @(0, 1605, 3010)   # 1605 = "not installed" (already gone) -> treat as non-fatal
$rebootNeeded = $false
$failures     = 0

foreach ($product in $targets) {
    $candidates = @(Get-UninstallCandidates -Product $product)
    if ($candidates.Count -eq 0) {
        Write-Log "No usable uninstall command for '$($product.DisplayName)'. Skipping." 'ERROR'
        $failures++
        continue
    }

    if (-not $Force) {
        $answer = Read-Host "Uninstall '$($product.DisplayName)'? [Y/N]"
        if ($answer -notmatch '^(y|yes)$') {
            Write-Log "Skipped by user: $($product.DisplayName)" 'WARN'
            continue
        }
    }

    if (-not $PSCmdlet.ShouldProcess($product.DisplayName, 'Uninstall')) {
        continue
    }

    $removed  = $false
    $lastCode = $null
    foreach ($cmd in $candidates) {
        Write-Log "Uninstalling '$($product.DisplayName)' via $($cmd.Kind): $($cmd.File) $($cmd.Args)"
        try {
            if ([string]::IsNullOrWhiteSpace($cmd.Args)) {
                $proc = Start-Process -FilePath $cmd.File -Wait -PassThru -WindowStyle Hidden
            }
            else {
                $proc = Start-Process -FilePath $cmd.File -ArgumentList $cmd.Args -Wait -PassThru -WindowStyle Hidden
            }
            $lastCode = $proc.ExitCode
            if ($successCodes -contains $lastCode) {
                if ($lastCode -eq 3010) { $rebootNeeded = $true }
                Write-Log "Removed '$($product.DisplayName)' (exit $lastCode via $($cmd.Kind))." 'OK'
                $removed = $true
                break
            }
            Write-Log "Method '$($cmd.Kind)' returned exit $lastCode; trying next method if available." 'WARN'
        }
        catch {
            Write-Log "Method '$($cmd.Kind)' threw: $($_.Exception.Message)" 'WARN'
        }
    }

    if (-not $removed) {
        Write-Log "All uninstall methods failed for '$($product.DisplayName)' (last exit $lastCode)." 'ERROR'
        Write-Log "  QuietUninstallString: $($product.QuietUninstallString)"
        Write-Log "  UninstallString:      $($product.UninstallString)"
        $failures++
    }
}

# --- Residual file cleanup ------------------------------------------------
function Remove-ResidualFiles {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string[]]$Paths)

    $removed = 0
    foreach ($path in $Paths) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        if (-not (Test-Path -LiteralPath $path)) { continue }

        # Hard safety guard: only delete paths under an Autodesk tree that
        # reference Revit/RVT 2026. Anything else is refused outright.
        if ($path -notmatch '(?i)\\Autodesk\\' -or $path -notmatch '(?i)Revit|RVT') {
            Write-Log "Refusing to remove unexpected residual path: $path" 'WARN'
            continue
        }
        if ($path -notmatch '2026') {
            Write-Log "Refusing to remove non-2026 residual path: $path" 'WARN'
            continue
        }

        if (-not $Force -and -not $WhatIfPreference) {
            $answer = Read-Host "Delete residual folder '$path'? [Y/N]"
            if ($answer -notmatch '^(y|yes)$') {
                Write-Log "Kept residual folder: $path" 'WARN'
                continue
            }
        }

        if ($PSCmdlet.ShouldProcess($path, 'Remove residual folder')) {
            try {
                Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
                Write-Log "Deleted residual folder: $path" 'OK'
                $removed++
            }
            catch {
                Write-Log "Failed to delete '$path': $($_.Exception.Message)" 'ERROR'
            }
        }
    }
    return $removed
}

if ($RemoveResidualFiles -and $failures -gt 0) {
    Write-Log 'Skipping residual cleanup because one or more products failed to uninstall. Re-run after the uninstall succeeds.' 'WARN'
}
elseif ($RemoveResidualFiles) {
    Write-Log 'Scanning for Revit 2026 residual folders...'
    $null = Remove-ResidualFiles -Paths $ResidualPaths
}

# --- Summary --------------------------------------------------------------
Write-Log '---------------------------------------------'
if ($failures -eq 0) {
    Write-Log "Completed. Shared Autodesk components were preserved." 'OK'
    if ($rebootNeeded) { Write-Log 'A reboot is recommended to finalize removal.' 'WARN' }
}
else {
    Write-Log "Completed with $failures failure(s). Review the log: $LogPath" 'ERROR'
}
Write-Log "Log saved to: $LogPath"

try { Stop-Transcript | Out-Null } catch { }

if ($failures -gt 0) { exit 3 }
elseif ($rebootNeeded) { exit 3010 }
else { exit 0 }
