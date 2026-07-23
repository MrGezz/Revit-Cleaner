<#
.SYNOPSIS
    Uninstalls an Autodesk Revit product (any year) and its orphaned add-ins on
    Windows, preserving shared and cross-version Autodesk components.

.DESCRIPTION
    Discovers the installed Autodesk Revit product for the target year
    (-ProductYear, default 2026) by reading the Windows "Uninstall" registry
    hives (64-bit, 32-bit/WOW6432Node, and per-user), then runs each
    vendor-registered uninstaller silently.

    Scope is deliberately conservative: the core Revit <year> application is
    always removed; with -IncludeAddins (default on) every product whose name
    references Revit AND the target year is also removed (add-ins, content
    packs, exporters, DB Link, IFC, interop tools). Shared and cross-version
    components (Material/Content Libraries, Licensing / Desktop Licensing
    Service, Genuine Service, Identity Manager, ODIS installer, Autodesk Access,
    RealDWG, Content Catalog year-range packs) are always preserved because
    other Autodesk products (Civil 3D, AutoCAD, Navisworks, other Revit years)
    depend on them.

    Resolution order for each product's uninstall command:
        1. msiexec /x {ProductCode} /qn /norestart   (when WindowsInstaller = 1)
        2. QuietUninstallString (vendor-provided silent command)
        3. UninstallString (run directly; --silent attempted for ODIS EXE
           uninstallers, with the exact vendor command kept as an auto fallback)

.PARAMETER ProductYear
    Four-digit Revit release year to target (e.g. 2023, 2024, 2025, 2026).
    Default: 2026. Everything - the core product match, the orphaned-add-in
    sweep, the residual folders, the residual path guard, and the self-elevation
    relaunch - is scoped to this year.

.PARAMETER IncludeAddins
    Also remove every product whose name references Revit and the target year
    - add-ins, content packs, exporters, DB Link, IFC, interop tools, etc. -
    which are orphaned once the core application is gone. Default: $true.
    Disable with -IncludeAddins:$false. Cross-version and shared components
    (Content Catalog year-range packs, RealDWG, version-neutral interop
    managers, material libraries, licensing) are always preserved.

.PARAMETER RemoveResidualFiles
    After a successful uninstall, delete leftover Revit <year>-specific folders
    (per-user settings/journals, add-in manifests, RVT content/templates, and
    any residual Revit <year> program folder). Default: $true. Disable with
    -RemoveResidualFiles:$false. A runtime guard only permits deletion of paths
    under an Autodesk tree that reference Revit/RVT and the target year; shared
    Autodesk trees are never removed.

.PARAMETER NeutralizeBrokenCustomActions
    When an MSI uninstall attempt exits 1603 and its verbose log shows
    "Internal Error 2753" (a custom action sourced from an installed file whose
    component registration is damaged), copy the cached package from
    C:\Windows\Installer to %TEMP%, condition the named action out ('0' =
    never run) in the COPY, and retry the uninstall from the patched copy.
    The protected cache is never modified (recent Windows builds refuse writes
    there even elevated). Surgical alternative to the Microsoft Program
    Install and Uninstall Troubleshooter: the rest of the uninstall still runs
    normally with full component cleanup and rollback.
    Default: $true. Disable with -NeutralizeBrokenCustomActions:$false.

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
    %TEMP%\Uninstall-Revit<year>_<timestamp>.log

.EXAMPLE
    # Preview what would be removed for the default year (2026), no changes:
    powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit.ps1 -ListOnly

.EXAMPLE
    # Uninstall Revit 2024 + its add-ins + residual files, prompting each step:
    powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit.ps1 -ProductYear 2024

.EXAMPLE
    # Fully unattended and silent for Revit 2025, close Revit if open:
    powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit.ps1 -ProductYear 2025 -StopRevit -Force

.EXAMPLE
    # Core product only (skip add-ins and residual cleanup):
    powershell -ExecutionPolicy Bypass -File .\Uninstall-Revit.ps1 -ProductYear 2026 -IncludeAddins:$false -RemoveResidualFiles:$false

.NOTES
    Requires an elevated (Administrator) session; the script self-elevates.
    Exit code 0 = success, 3010 = success (reboot required), 3 = partial failure,
    2 = nothing found, 1 = aborted.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [ValidatePattern('^\d{4}$')]
    [string]$ProductYear       = '2026',
    [bool]$IncludeAddins       = $true,
    [bool]$IncludeMaterialLibraries = $false,
    [bool]$RemoveResidualFiles = $true,
    [bool]$NeutralizeBrokenCustomActions = $true,
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
# $ProductYear is supplied by the -ProductYear parameter (default 2026) and
# scopes every match below.

# Core product: always targeted.
$CorePatterns = @(
    "Autodesk Revit $ProductYear"
)

# Shared / cross-product components: NEVER touched in this mode.
$SharedExclusions = @(
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

# Lock shared material packages out dynamically if explicit sweep is off
if (-not $IncludeMaterialLibraries) {
    $SharedExclusions += '*Material Library*'
    $SharedExclusions += '*Advanced Material Library*'
}

# Revit <year>-specific residual folders removed only with -RemoveResidualFiles.
# Every entry references a Revit/RVT <year> path; shared Autodesk trees are never
# listed here, and a runtime guard re-verifies each path before deletion.
$ResidualPaths = @(
    (Join-Path $env:APPDATA        "Autodesk\Revit\Autodesk Revit $ProductYear"),
    (Join-Path $env:LOCALAPPDATA   "Autodesk\Revit\Autodesk Revit $ProductYear"),
    (Join-Path $env:APPDATA        "Autodesk\Revit\Addins\$ProductYear"),
    (Join-Path $env:ProgramData    "Autodesk\Revit\Addins\$ProductYear"),
    (Join-Path $env:ProgramData    "Autodesk\RVT $ProductYear"),
    (Join-Path ${env:ProgramFiles} "Autodesk\Revit $ProductYear")
)

# Directory-property override shared by the MSI-PropsOverride attempt and the
# 2753 patched-copy retry. INSTALLDIR must CONTAIN the "Revit <year>" token so
# the MSI's Type-51 DIRCA_INSTALLDIR action (condition:
# NOT INSTALLDIR><ADSK_INSTALL_PATH) skips instead of rebuilding INSTALLDIR as
# the bare relative fragment "Revit <year>\" that kills CostFinalize with
# 1314/1606. NO trailing backslash before the closing quote (\" escapes the
# quote and mangles every argument after it); ROOTDRIVE stays unquoted (no
# spaces) because its value MUST end in a backslash.
$FallbackDirProps = ' ROOTDRIVE=C:\ INSTALLDIR="' + ${env:ProgramFiles} + '\Autodesk\Revit ' + $ProductYear + '"'

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
    $passArgs += ('-ProductYear {0}'         -f $ProductYear)
    $passArgs += ('-IncludeAddins:${0}'      -f $IncludeAddins)
    $passArgs += ('-IncludeMaterialLibraries:${0}' -f $IncludeMaterialLibraries)
    $passArgs += ('-RemoveResidualFiles:${0}' -f $RemoveResidualFiles)
    $passArgs += ('-NeutralizeBrokenCustomActions:${0}' -f $NeutralizeBrokenCustomActions)
    if ($StopRevit) { $passArgs += '-StopRevit' }
    if ($ListOnly)  { $passArgs += '-ListOnly' }
    if ($Force)     { $passArgs += '-Force' }
    if ($LogPath)   { $passArgs += ("-LogPath '{0}'" -f ($LogPath -replace "'", "''")) }

    # powershell.exe -File CANNOT bind [bool] parameters (nor -Switch:$false):
    # PS 5.1 passes every -File argument as a literal string and rejects it with
    # "Boolean parameters accept only Boolean values and numbers" (verified on
    # this machine for True / False / 1 / 0 / $false). The elevated child would
    # die at parameter binding BEFORE Start-Transcript - no log, elevation
    # "does nothing". Relaunch through -Command with a single-quoted
    # call-operator path instead: $true/$false literals then parse natively and
    # a spaced script path still survives as one argument. The trailing
    # "; exit $LASTEXITCODE" is required: in -Command mode an "exit N" inside
    # the invoked SCRIPT only sets $LASTEXITCODE - without re-exiting, the
    # child process collapses every non-zero script exit to 1 (verified:
    # exit 42 came back as 1 without it, 42 with it), destroying the
    # 0/2/3/3010 exit-code contract relayed by the non-elevated parent.
    $qPath   = $PSCommandPath -replace "'", "''"
    $cmdLine = '-NoProfile -ExecutionPolicy Bypass -Command "& ''{0}'' {1}; exit $LASTEXITCODE"' -f $qPath, ($passArgs -join ' ')

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
    $LogPath = Join-Path $env:TEMP "Uninstall-Revit${ProductYear}_$stamp.log"
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
                InstallLocation      = Get-Prop $props 'InstallLocation'
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
                       ($name -match '(?i)\bRevit\b|\bRVT\b|\bNavisworks\b.*?\bExport') -and
                       ($name -like "*$ProductYear*")
 
        $isMaterialYear = $IncludeMaterialLibraries -and 
                          ($name -match '(?i)Material Library') -and 
                          ($name -like "*$ProductYear*")
 
        $isCore -or $isRevitYear -or $isMaterialYear
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
            Write-Log "No Revit $ProductYear residual folders found."
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


# Resolves the locally CACHED .msi for a given ProductCode via the Windows
# Installer COM API (MSI "LocalPackage" property). Uninstalling from this
# literal file path bypasses SourceList/network-source resolution entirely -
# the fix for "Error 1606: Could not access network location <share>\" during
# /x, which the caller otherwise only sees wrapped as a generic exit 1603
# (visible instead in the per-product MSI*.LOG next to %TEMP%).
function Get-MsiLocalPackage {
    param([string]$ProductCode)
    # Pre-initialize: if New-Object throws, the finally block would otherwise
    # reference an undefined variable and itself throw under StrictMode Latest.
    $installer = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        # INSTALLPROPERTY_LOCALPACKAGE = "LocalPackage"
        $local = $installer.ProductInfo($ProductCode, 'LocalPackage')
        if (-not [string]::IsNullOrWhiteSpace($local) -and (Test-Path -LiteralPath $local)) {
            return $local
        }
    }
    catch {
        Write-Log "LocalPackage lookup failed for $ProductCode`: $($_.Exception.Message)" 'WARN'
    }
    finally {
        if ($installer) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($installer) }
    }
    return $null
}

# Squished (registry-form) GUID for a ProductCode, e.g.
# {7346B4A0-2300-0510-0000-705C0D862004} -> 0A4B6437...  (block-reversed hex,
# no braces/dashes). Windows Installer keys its UserData/Installer\Products
# registrations under this form, not the plain ProductCode.
function Get-MsiSquishedGuid {
    param([string]$ProductCode)
    if ($ProductCode -notmatch '^\{([0-9a-fA-F\-]{36})\}$') { return $null }
    $raw = $matches[1] -replace '-'
    $parts = @(
        $raw.Substring(0,8), $raw.Substring(8,4), $raw.Substring(12,4),
        $raw.Substring(16,2), $raw.Substring(18,2), $raw.Substring(20,2),
        $raw.Substring(22,2), $raw.Substring(24,2), $raw.Substring(26,2),
        $raw.Substring(28,2), $raw.Substring(30,2)
    )
    return ($parts | ForEach-Object { $c = $_.ToCharArray(); [array]::Reverse($c); -join $c }) -join ''
}

# Diagnostic-only: dump every SourceList value that could be feeding the
# 1606 "Revit 2023\" concatenation, so the log shows the ACTUAL bad entry
# instead of just the symptom. Run this BEFORE the purge below and compare.
function Show-MsiSourceListDump {
    param([string]$ProductCode)
    $squished = Get-MsiSquishedGuid -ProductCode $ProductCode
    if (-not $squished) { return }
    $roots = @(
        "HKLM:\SOFTWARE\Classes\Installer\Products\$squished\SourceList",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$squished\SourceList"
    )
    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        Write-Log "SourceList dump: $root" 'WARN'
        try {
            $props = Get-ItemProperty -LiteralPath $root -ErrorAction SilentlyContinue
            foreach ($n in @('PackageName','LastUsedSource')) {
                $v = Get-Prop $props $n
                if ($v) { Write-Log "    $n = $v" 'WARN' }
            }
        } catch { }
        foreach ($sub in 'Net','URL','Media') {
            $subPath = Join-Path $root $sub
            if (-not (Test-Path -LiteralPath $subPath)) { continue }
            try {
                $sp = Get-ItemProperty -LiteralPath $subPath -ErrorAction SilentlyContinue
                $sp.PSObject.Properties |
                    Where-Object { $_.Name -notmatch '^PS' } |
                    ForEach-Object { Write-Log "    $sub\$($_.Name) = $($_.Value)" 'WARN' }
            } catch { }
        }
    }
}

# THE ACTUAL FIX for "Error 1606: Could not access network location <x>\":
# directly purge and rewrite the product's SourceList registry entries so
# msiexec has nothing dangling left to concatenate against. This is the
# standard, Microsoft-documented remediation for a stale/relative SourceList
# entry (post-migration, moved/deleted deployment shares) - registry-level,
# not COM, so it can't silently no-op the way SourceListClearAll's late-bound
# InvokeMember call can. Also rewrites PackageName to the LOCAL cached .msi
# basename (when known) so any later PackageName+LastUsedSource concatenation
# resolves to a real local file instead of the dead "Revit 2023\" fragment.
function Clear-MsiSourceListRegistry {
    param([string]$ProductCode, [string]$LocalPackagePath)
    $squished = Get-MsiSquishedGuid -ProductCode $ProductCode
    if (-not $squished) { return }
    $roots = @(
        "HKLM:\SOFTWARE\Classes\Installer\Products\$squished\SourceList",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$squished\SourceList"
    )
    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        foreach ($sub in 'Net','URL') {
            $subPath = Join-Path $root $sub
            if (Test-Path -LiteralPath $subPath) {
                try {
                    Remove-Item -LiteralPath $subPath -Recurse -Force -ErrorAction Stop
                    Write-Log "Removed stale $sub entries under $root" 'OK'
                } catch {
                    Write-Log "Could not remove $subPath : $($_.Exception.Message)" 'WARN'
                }
            }
        }
        try {
            if (Get-ItemProperty -LiteralPath $root -Name 'LastUsedSource' -ErrorAction SilentlyContinue) {
                Remove-ItemProperty -LiteralPath $root -Name 'LastUsedSource' -ErrorAction SilentlyContinue
                Write-Log "Cleared LastUsedSource under $root" 'OK'
            }
        } catch { }
        if ($LocalPackagePath -and (Test-Path -LiteralPath $LocalPackagePath)) {
            try {
                Set-ItemProperty -LiteralPath $root -Name 'PackageName' `
                    -Value (Split-Path -Leaf $LocalPackagePath) -ErrorAction Stop
                Write-Log "Rewrote PackageName -> $(Split-Path -Leaf $LocalPackagePath) under $root" 'OK'
            } catch {
                Write-Log "Could not rewrite PackageName under $root : $($_.Exception.Message)" 'WARN'
            }
        }
    }
}



function Repair-MsiUserDataCache {
    param([string]$ProductCode, [string]$TargetYear)
    # Deep-patch Error 1606 by recalculating to a "Squished" internal MSI GUID cache
    $squished = Get-MsiSquishedGuid -ProductCode $ProductCode
    if (-not $squished) { return }

    $msiHive = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$squished\InstallProperties"
    
    if (Test-Path -LiteralPath $msiHive) {
        try {
            $il = Get-ItemProperty -Path $msiHive -Name 'InstallLocation' -ErrorAction SilentlyContinue
            if ($null -eq $il -or [string]::IsNullOrWhiteSpace($il.InstallLocation) -or $il.InstallLocation -notmatch '^([a-zA-Z]:[\\/]|\\\\)') {
                $fixPath = Join-Path ${env:ProgramFiles} "Autodesk\Revit $TargetYear\"
                Set-ItemProperty -LiteralPath $msiHive -Name 'InstallLocation' -Value $fixPath -ErrorAction Stop
                Write-Log "Patched internal MSI cache location to bypass Error 1606:`n      New: $fixPath" 'WARN'
            }
        } catch { 
            Write-Log "Failed patching native cache InstallLocation: $($_.Exception.Message)" 'WARN' 
        }
    }
}


# Surgical Error-2753 remediation. "Internal Error 2753. <key>" means a custom
# action sourced from an INSTALLED FILE (CustomAction base type 17/18/21/22)
# whose component registration is damaged - Windows Installer cannot resolve
# the file, and the whole uninstall aborts even though every other action would
# succeed. The cached package in C:\Windows\Installer CANNOT be edited in
# place: current Windows builds ACL it against writes even from an elevated
# admin (observed: transact OpenDatabase fails on build 26200). So this copies
# the cached .msi to %TEMP%, sets the named action's InstallExecuteSequence
# Condition to '0' (= never run) in the COPY, and returns the patched copy's
# path - the caller retries the uninstall FROM that path. Path-based
# maintenance is accepted because the PackageCode is unchanged, and the rest
# of the uninstall runs normally with full component cleanup and rollback,
# unlike the Microsoft troubleshooter's force-removal, which just rips the
# registration. On repeat calls pass -PatchedMsi so later neutralizations
# accumulate in the SAME copy instead of starting over from the cache.
# Returns the patched copy's path, or $null if there was nothing to do.
function Repair-MsiBrokenCustomAction {
    param([string]$ProductCode, [string]$VerboseLog, [string]$PatchedMsi)

    if ([string]::IsNullOrWhiteSpace($VerboseLog) -or -not (Test-Path -LiteralPath $VerboseLog)) { return $null }
    if (-not (Select-String -Path $VerboseLog -Pattern 'Error 2753' -Quiet)) { return $null }

    # The action that aborted: first non-INSTALL "Return value 3" line.
    $fail = Select-String -Path $VerboseLog -Pattern 'Action ended .*?: (.+)\. Return value 3\.' |
        Where-Object { $_.Matches[0].Groups[1].Value -ne 'INSTALL' } |
        Select-Object -First 1
    if (-not $fail) { return $null }
    $action = $fail.Matches[0].Groups[1].Value
    if ($action -match "'") { return $null }   # never build WI-SQL from an apostrophed name

    if (-not [string]::IsNullOrWhiteSpace($PatchedMsi) -and (Test-Path -LiteralPath $PatchedMsi)) {
        # Accumulate into the existing patched copy.
        $target = $PatchedMsi
    }
    else {
        $localMsi = Get-MsiLocalPackage -ProductCode $ProductCode
        if (-not $localMsi) {
            Write-Log "2753 on '$action' but no cached package found for $ProductCode - cannot neutralize." 'WARN'
            return $null
        }
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        # The patched copy MUST keep the registered package's exact file name:
        # during the /fv repair, source resolution probes
        # SOURCEDIR + <registered PackageName> (observed: it looked for
        # %TEMP%\1d05a2.msi and failed 2203/1316 'Error determining package
        # source type' because our copy was named 1d05a2_patched_*.msi).
        # A per-run subfolder keeps the name without colliding in %TEMP%.
        $patchDir = Join-Path $env:TEMP ('RevitCleanerPatch_' + $stamp)
        $target   = Join-Path $patchDir ([IO.Path]::GetFileName($localMsi))
        # Pristine snapshot of the registered cached package: the caller's
        # recache step (/fv) will REPLACE the cache with the patched copy, so
        # keep an unmodified original for manual rollback.
        $pristine = Join-Path $env:TEMP ('{0}_pristine_{1}.msi' -f `
            [IO.Path]::GetFileNameWithoutExtension($localMsi), $stamp)
        try {
            if (-not (Test-Path -LiteralPath $patchDir)) {
                $null = New-Item -ItemType Directory -Path $patchDir -Force -ErrorAction Stop
            }
            Copy-Item -LiteralPath $localMsi -Destination $pristine -Force -ErrorAction Stop
            Copy-Item -LiteralPath $localMsi -Destination $target -Force -ErrorAction Stop
            Write-Log "Pristine cached-package backup: $pristine" 'INFO'
        }
        catch {
            Write-Log "Cannot copy cached package to %TEMP% ($($_.Exception.Message)) - cannot neutralize." 'ERROR'
            return $null
        }
    }

    $installer = $null; $db = $null; $view1 = $null; $view2 = $null; $rec = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        # 1 = msiOpenDatabaseModeTransact. DIRECT dispatch: reflection
        # InvokeMember on Installer.OpenDatabase throws DISP_E_TYPEMISMATCH on
        # this PS 5.1 host (observed live: direct $installer.OpenDatabase()
        # succeeds where InvokeMember fails on the identical arguments), so
        # try direct first and keep InvokeMember only as a fallback.
        $db = $null
        try { $db = $installer.OpenDatabase($target, 1) } catch { }
        if ($null -eq $db) {
            $db = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($target, 1))
        }
        # EVERY void COM call below is [void]-cast: direct-dispatch void
        # methods emit $null into the function's OUTPUT STREAM in PS 5.1, so
        # without the casts this function returned an ARRAY (nulls + path).
        # String-interpolating that array space-joined it, msiexec received
        # /x "    C:\...patched.msi", resolved it as RELATIVE (CWD prepended,
        # note 2203) and failed 1619. Observed live; never return bare COM
        # call results from an advanced function.
        $view1 = $db.OpenView("SELECT ``Action`` FROM ``InstallExecuteSequence`` WHERE ``Action`` = '$action'")
        [void]$view1.Execute()
        $rec = $view1.Fetch()
        if ($null -eq $rec) {
            Write-Log "Action '$action' not present in InstallExecuteSequence - cannot neutralize." 'WARN'
            return $null
        }
        [void]$view1.Close()
        $view2 = $db.OpenView("UPDATE ``InstallExecuteSequence`` SET ``Condition`` = '0' WHERE ``Action`` = '$action'")
        [void]$view2.Execute()
        [void]$view2.Close()
        [void]$db.Commit()
        Write-Log "Neutralized broken custom action '$action' (Error 2753) in patched copy: $target" 'OK'
        return $target
    }
    catch {
        Write-Log "Failed to neutralize '$action' in ${target}: $($_.Exception.Message)" 'ERROR'
        return $null
    }
    finally {
        # Release EVERY wrapper - including the SELECT view and its fetched
        # Record - then force a GC. Any un-finalized RCW keeps the .msi handle
        # open in THIS process, and the immediately-following msiexec then
        # fails 1619 (note 2203 / 0x80030020 STG_E_SHAREVIOLATION). Verified
        # live: exclusive reopen fails before this block's GC, succeeds after.
        foreach ($o in @($rec, $view1, $view2, $db, $installer)) {
            if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    }
}


# Clears a stale SourceList (the actual trigger for the 1606 network-location
# error) so any LATER msiexec call against the bare ProductCode - e.g. a
# retry, or another tool - no longer tries to reach the dead network share.
# Best-effort; failure here never blocks the uninstall itself.
function Clear-MsiSourceList {
    param([string]$ProductCode)
    # Pre-initialize for the finally block - see Get-MsiLocalPackage.
    $installer = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        # MsiSourceListClearAllEx via the Installer.SourceListClearAll method
        # (context 7 = MSIINSTALLCONTEXT_ALL, options 4 = MSICODE_PRODUCT).
        $installer.GetType().InvokeMember(
            'SourceListClearAll', 'InvokeMethod', $null, $installer,
            @($ProductCode, '', 7)) | Out-Null
        Write-Log "Cleared stale SourceList for $ProductCode" 'INFO'
    }
    catch {
        # Non-fatal - not every Windows Installer version exposes this verb,
        # and late-bound automation calls can fail to marshal silently.
        # Clear-MsiSourceListRegistry (called separately, unconditionally,
        # below) is the guaranteed path and does not depend on this succeeding.
    }
    finally {
        if ($installer) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($installer) }
    }
}


# Build an ordered list of uninstall attempts. If one method fails the caller
# falls through to the next.
function Get-UninstallCandidates {
    param($Product)
    $list = @()

    # 1) MSI, preferring the LOCALLY CACHED .msi over the bare product code.
    #    Uninstalling by ProductCode makes Windows Installer re-validate the
    #    ORIGINAL install source (SourceList) before it will proceed; if that
    #    was a network share that's now gone, MSI aborts with
    #    "Error 1606: Could not access network location ...", surfaced to the
    #    caller only as a generic exit 1603. Uninstalling from the cached
    #    LocalPackage path skips that source-resolution step entirely.
    if ($Product.WindowsInstaller -eq 1 -and $Product.KeyName -match '^\{[0-9A-Fa-f\-]{36}\}$') {
        
        # LAST-RESORT ONLY property override - do NOT inject this into the
        # primary attempts. Forcing INSTALLDIR/ROOTDRIVE on an uninstall of an
        # already-registered product can change a component/directory
        # CONDITION evaluation from what it was at install time, knocking a
        # component (and any File its CustomAction/Binary table references)
        # out of the action sequence - which is precisely what
        # "Error 2753: The File '[2]' is not marked for installation" means.
        # Kept available as a distinct, clearly-separate candidate at the
        # BOTTOM of the list so we still try it, but only after every
        # unmodified attempt has failed.
        # Shared with the 2753 patched-copy retry - defined once in the
        # configuration section (see $FallbackDirProps comment there for why
        # INSTALLDIR must contain the "Revit <year>" token and how the
        # quoting works).
        $fallBackProps = $FallbackDirProps

        # Verbose MSI log (separate from the terse system-default MSI*.LOG in
        # %TEMP%) - the terse log gives us the error CODE only; this gives us
        # the actual Action/Component/File sequence around the failure, which
        # is what we need to diagnose 2753 with certainty instead of guessing.
        # One verbose log PER ATTEMPT: /L*V truncates on open, so a shared
        # filename means the second attempt silently overwrites the first
        # attempt's evidence (observed: LocalPackage and bare-MSI attempts of
        # the same run wrote the identical file).
        $vlogBase = Join-Path $env:TEMP ("MSIVerbose_{0}_{1}" -f `
            ($Product.KeyName -replace '[{}]',''), (Get-Date -Format 'yyyyMMdd_HHmmss'))
        $vlogMsiLocalPath = $vlogBase + '_LocalPackage.log'
        $vlogMsiPath      = $vlogBase + '_MSI.log'
        $vlogMsiPropsPath = $vlogBase + '_PropsOverride.log'
        $vlogMsiLocal = ' /L*V "' + $vlogMsiLocalPath + '"'
        $vlogMsi      = ' /L*V "' + $vlogMsiPath + '"'
        $vlogMsiProps = ' /L*V "' + $vlogMsiPropsPath + '"'
 
        $localMsi = Get-MsiLocalPackage -ProductCode $Product.KeyName

        # UNCONDITIONAL - the 1606 concatenation happens off the ProductCode's
        # registered SourceList regardless of whether /x targets the bare
        # ProductCode or the LocalPackage file path, so this must run before
        # EVERY attempt, not only when LocalPackage lookup fails.
        Show-MsiSourceListDump -ProductCode $Product.KeyName
        Clear-MsiSourceList -ProductCode $Product.KeyName
        Clear-MsiSourceListRegistry -ProductCode $Product.KeyName -LocalPackagePath $localMsi

        if ($localMsi) {
            $list += [pscustomobject]@{ File = 'msiexec.exe'; Args = "/x `"$localMsi`" /qn /norestart$vlogMsiLocal"; Kind = 'MSI-LocalPackage'; Vlog = $vlogMsiLocalPath }
        }

        $list += [pscustomobject]@{ File = 'msiexec.exe'; Args = "/x $($Product.KeyName) /qn /norestart$vlogMsi"; Kind = 'MSI'; Vlog = $vlogMsiPath }
        # Property-override attempt, LAST - see comment above.
        $list += [pscustomobject]@{ File = 'msiexec.exe'; Args = "/x $($Product.KeyName) /qn /norestart$fallBackProps$vlogMsiProps"; Kind = 'MSI-PropsOverride'; Vlog = $vlogMsiPropsPath }
    }

    # 2) Vendor-provided silent command.
    if (-not [string]::IsNullOrWhiteSpace($Product.QuietUninstallString)) {
        $s = Split-Command $Product.QuietUninstallString
        $list += [pscustomobject]@{ File = $s.File; Args = $s.Args; Kind = 'Quiet'; Vlog = $null }
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
            $list += [pscustomobject]@{ File = $s.File; Args = $s.Args; Kind = 'Fallback'; Vlog = $null }
        }
        else {
            $s = Split-Command $raw
            if ($s.Args -notmatch '(?i)(^|\s)(--silent|-silent|/silent|-q|/q)(\s|$)') {
                $list += [pscustomobject]@{ File = $s.File; Args = ($s.Args + ' --silent').Trim(); Kind = 'Silent'; Vlog = $null }
            }
            $list += [pscustomobject]@{ File = $s.File; Args = $s.Args; Kind = 'Fallback'; Vlog = $null }
        }
    }

    # De-duplicate, collapsing msiexec variants that target the same GUID
    # (e.g. "msiexec.exe /x {g}" and "MsiExec.exe /X{g}") into a single attempt.
    $seen = @{}
    $unique = @()
    foreach ($c in $list) {
        $norm = ('{0} {1}' -f $c.File, $c.Args).ToLowerInvariant()
        $guid = [regex]::Match($norm, '\{[0-9a-f\-]{36}\}')
        # GUID-collapse ONLY plain "/x {guid}" invocations. A candidate that
        # carries PROPERTY=value overrides is functionally different: keying it
        # by GUID alone silently deleted MSI-PropsOverride from the attempt list
        # (transcripts show only two methods ever ran), so the 1606 workaround
        # never fired.
        if ($norm -match 'msiexec' -and $guid.Success -and $norm -notmatch '\s\w+=') { $key = 'msi:' + $guid.Value }
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

    # Consent first, registry surgery second. Repair-MsiUserDataCache and the
    # SourceList dump/purge inside Get-UninstallCandidates all WRITE to the
    # machine; running them before Read-Host/ShouldProcess mutated state even
    # when the answer was No or the run was -WhatIf.
    if ($product.WindowsInstaller -eq 1) {
        Repair-MsiUserDataCache -ProductCode $product.KeyName -TargetYear $ProductYear
    }

    $candidates = @(Get-UninstallCandidates -Product $product)
    if ($candidates.Count -eq 0) {
        Write-Log "No usable uninstall command for '$($product.DisplayName)'. Skipping." 'ERROR'
        $failures++
        continue
    }

    $removed  = $false
    $lastCode = $null
    foreach ($cmd in $candidates) {
        # Bounded retry loop per method: a 1603 whose verbose log shows
        # Internal Error 2753 gets the broken custom action neutralized in a
        # patched %TEMP% copy of the cached package, and the uninstall is
        # retried FROM that copy (with the directory-property override, since
        # the copy still carries DIRCA_INSTALLDIR). Each attempt can only
        # surface one broken file-sourced action at a time, hence the loop;
        # 5 is far above anything seen in practice (this package carries
        # exactly one file-sourced custom action).
        $caRepairs  = 0
        $patchedMsi = $null
        while ($true) {
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
                if ($lastCode -eq 1603 -and $NeutralizeBrokenCustomActions -and $caRepairs -lt 5 -and
                    $cmd.Kind -like 'MSI*' -and $cmd.Vlog) {
                    $patch = Repair-MsiBrokenCustomAction -ProductCode $product.KeyName `
                        -VerboseLog $cmd.Vlog -PatchedMsi $patchedMsi
                    # Defense in depth against output-stream pollution: keep
                    # only the last non-empty string (the returned path).
                    $patch = @($patch | Where-Object { $_ -is [string] -and -not [string]::IsNullOrWhiteSpace($_) }) |
                        Select-Object -Last 1
                    if ($patch) {
                        $caRepairs++
                        $patchedMsi = $patch
                        # Maintenance mode IGNORES the package path argument's
                        # tables: "Package we're running from ==>" is always
                        # the REGISTERED cache (verified in the verbose logs),
                        # so /x "<patched>.msi" never sees the neutralized
                        # condition. The supported route is a RECACHE repair -
                        # msiexec /fv <patched> replaces the registered cached
                        # package with ours (PackageCode unchanged, and the
                        # registered language transform does not touch the
                        # neutralized row - verified via ApplyTransform) -
                        # then a normal product-code /x runs from it. The /fv
                        # carries the same directory-property override because
                        # repair costing hits DIRCA_INSTALLDIR too.
                        $recacheVlog = $cmd.Vlog -replace '\.log$', ('_recache{0}.log' -f $caRepairs)
                        $recacheArgs = "/fv `"$patchedMsi`" /qn /norestart$FallbackDirProps" + ' /L*V "' + $recacheVlog + '"'
                        Write-Log "Recaching patched package: msiexec.exe $recacheArgs"
                        $rc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $recacheArgs -Wait -PassThru -WindowStyle Hidden
                        if ($rc.ExitCode -notin 0, 3010) {
                            Write-Log "Recache (/fv) failed with exit $($rc.ExitCode) - cannot apply the 2753 fix. See $recacheVlog" 'ERROR'
                            break
                        }
                        if ($rc.ExitCode -eq 3010) { $rebootNeeded = $true }
                        Write-Log "Recache OK - registered cache now carries the neutralized action." 'OK'
                        $retryVlog = $cmd.Vlog -replace '\.log$', ('_recached{0}.log' -f $caRepairs)
                        $cmd = [pscustomobject]@{
                            File = 'msiexec.exe'
                            Args = "/x $($product.KeyName) /qn /norestart$FallbackDirProps" + ' /L*V "' + $retryVlog + '"'
                            Kind = 'MSI-Recached'
                            Vlog = $retryVlog
                        }
                        Write-Log "Retrying by product code from the recached package (repair $caRepairs of 5)." 'WARN'
                        continue
                    }
                }
                Write-Log "Method '$($cmd.Kind)' returned exit $lastCode; trying next method if available." 'WARN'
            }
            catch {
                Write-Log "Method '$($cmd.Kind)' threw: $($_.Exception.Message)" 'WARN'
            }
            break
        }
        if ($removed) { break }
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
        # reference Revit/RVT and the target year. Anything else is refused.
        if ($path -notmatch '(?i)\\Autodesk\\' -or $path -notmatch '(?i)Revit|RVT') {
            Write-Log "Refusing to remove unexpected residual path: $path" 'WARN'
            continue
        }
        if ($path -notmatch $ProductYear) {
            Write-Log "Refusing to remove non-$ProductYear residual path: $path" 'WARN'
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
    Write-Log "Scanning for Revit $ProductYear residual folders..."
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