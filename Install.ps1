<#
.SYNOPSIS
    Install the lib_access_restapi modules into an existing Access database.

.DESCRIPTION
    Opens the supplied .accdb / .mdb via COM, adds the MSXML 6.0
    type-library reference (belt-and-suspenders; the library itself
    is fully late-bound via CreateObject), and imports every production
    VBA module from the script directory. Tests are skipped by default;
    pass -IncludeTests to import them too.

    Run from the directory containing the .bas / .cls files:

        powershell -ExecutionPolicy Bypass -File .\Install.ps1 -Database C:\path\to\app.accdb

.PARAMETER Database
    Full path to the target .accdb / .mdb. The file must already exist
    and must not be open in another Access instance.

.PARAMETER IncludeTests
    Also import clsResponseCollector.cls, clsLateBoundReceiver.cls and
    modRestTest.bas. Off by default.

.PARAMETER Force
    Replace existing modules with the same name. Without this flag,
    existing components are left untouched and a warning is printed.

.NOTES
    Prereqs:
      - Microsoft Access (64-bit recommended).
      - Trust Center -> Macro Settings -> "Trust access to the VBA
        project object model" must be enabled.
      - Target database must not already be open in Access.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Database,

    [switch] $IncludeTests,
    [switch] $Force,
    [switch] $Visible
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$ProductionModules = @(
    'modJson.bas',
    'IHttpCallback.cls',
    'clsHttpResponse.cls',
    'clsAsyncCall.cls',
    'clsHttpRequest.cls',
    'modRestPump.bas',
    'modRest.bas'
)

$TestModules = @(
    'clsResponseCollector.cls',
    'clsLateBoundReceiver.cls',
    'modRestTest.bas'
)

$Modules = $ProductionModules
if ($IncludeTests) { $Modules += $TestModules }

# MSXML 6.0 type library GUID. Stable across Windows versions; ships in OS.
$MsxmlGuid  = '{F5078F18-C551-11D3-89B9-0000F81FE221}'
$MsxmlMajor = 6
$MsxmlMinor = 0

function Release-Com([object]$obj) {
    if ($obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
    }
}

# Return a path to a CRLF-normalized copy of $srcPath. The VBE rejects
# LF-only .bas/.cls files with "Erwartet Anweisungsende" because the
# header (VERSION 1.0 CLASS / BEGIN / END / Attribute lines) is parsed
# expecting CRLF. Stage into a temp subfolder cleaned up at exit.
function Convert-ToCrlf([string]$srcPath, [string]$stageDir) {
    if (-not (Test-Path -LiteralPath $stageDir)) {
        New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
    }
    $dstPath = Join-Path $stageDir (Split-Path -Leaf $srcPath)
    $raw = [System.IO.File]::ReadAllText($srcPath)
    $normalized = ($raw -replace "`r`n", "`n") -replace "`n", "`r`n"
    [System.IO.File]::WriteAllText($dstPath, $normalized, (New-Object System.Text.UTF8Encoding $false))
    return $dstPath
}

function Get-ComponentName([string]$filePath) {
    # Component name in VBA == file basename in our convention. Could parse
    # the "Attribute VB_Name = "..."" line for robustness, but the
    # filename is the source of truth here.
    return [System.IO.Path]::GetFileNameWithoutExtension($filePath)
}

# --- Pre-flight checks ---

if (-not (Test-Path -LiteralPath $Database)) {
    Write-Error "Database file not found: $Database"
    exit 2
}

$Database = (Resolve-Path -LiteralPath $Database).Path

foreach ($m in $Modules) {
    $path = Join-Path $ScriptDir $m
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Error "Missing module file: $path"
        exit 2
    }
}

Write-Host "Target database: $Database" -ForegroundColor Cyan
Write-Host "Module set:      $(if ($IncludeTests) {'production + tests'} else {'production'})" -ForegroundColor Cyan

# --- Drive Access ---

$access = $null
try {
    Write-Host "Starting Access..." -ForegroundColor Cyan
    $access = New-Object -ComObject Access.Application
} catch {
    Write-Error "Could not start Access via COM. Is Microsoft Access installed?"
    exit 3
}

try {
    $access.Visible = [bool]$Visible
    try { $access.AutomationSecurity = 1 } catch {}  # msoAutomationSecurityLow
    try { $access.DoCmd.SetWarnings($false) } catch {}

    Write-Host "Opening database..." -ForegroundColor Cyan
    try {
        $access.OpenCurrentDatabase($Database)
    } catch {
        Write-Error @"
Failed to open $Database.

Common causes:
  - The file is already open in another Access window. Close it and retry.
  - The path or file is wrong / corrupt.
  - The file is read-only or locked by another process.

Underlying error: $($_.Exception.Message)
"@
        exit 4
    }

    try {
        $vbe = $access.VBE
        $vbproj = $vbe.VBProjects.Item(1)
    } catch {
        Write-Error @"
Cannot reach the VBA project model.

In Access, open File -> Options -> Trust Center -> Trust Center Settings ->
Macro Settings, and tick "Trust access to the VBA project object model".
Restart Access and run this script again.
"@
        exit 5
    }

    # --- MSXML 6.0 reference ---

    Write-Host "Adding MSXML 6.0 reference..." -ForegroundColor Cyan
    try {
        $vbproj.References.AddFromGuid($MsxmlGuid, $MsxmlMajor, $MsxmlMinor) | Out-Null
        Write-Host "  added" -ForegroundColor Green
    } catch {
        if ($_.Exception.Message -match 'already' -or
            $_.Exception.Message -match 'Name conflicts') {
            Write-Host "  already present, skipped" -ForegroundColor DarkGray
        } else {
            throw
        }
    }

    # --- Module imports ---

    Write-Host "Importing modules..." -ForegroundColor Cyan
    $imported = 0
    $skipped  = 0
    $replaced = 0
    $stamp    = [Guid]::NewGuid().ToString('N')
    $StageDir = Join-Path $ScriptDir "_stage_$stamp"

    foreach ($m in $Modules) {
        $src  = Join-Path $ScriptDir $m
        $name = Get-ComponentName $src

        $existing = $null
        try { $existing = $vbproj.VBComponents.Item($name) } catch { }

        $crlfPath = Convert-ToCrlf $src $StageDir

        if ($existing) {
            if ($Force) {
                Write-Host "  ~ $m (replacing existing $name)" -ForegroundColor Yellow
                $vbproj.VBComponents.Remove($existing) | Out-Null
                $vbproj.VBComponents.Import($crlfPath) | Out-Null
                $replaced++
            } else {
                Write-Host "  - $m (skipped, $name already exists; use -Force to replace)" -ForegroundColor DarkGray
                $skipped++
                continue
            }
        } else {
            Write-Host "  + $m" -ForegroundColor Green
            $vbproj.VBComponents.Import($crlfPath) | Out-Null
            $imported++
        }
    }

    # Skip an explicit compile-and-save: DoCmd.RunCommand(126) isn't
    # consistent across Access versions and on some setups triggers the
    # "Export Module As Text" dialog. CloseCurrentDatabase below persists
    # the imported modules; use Debug -> Compile in Access for a manual
    # compile check.

    Write-Host ""
    Write-Host "Done." -ForegroundColor Green
    Write-Host "  imported: $imported" -ForegroundColor Green
    if ($replaced -gt 0) { Write-Host "  replaced: $replaced" -ForegroundColor Yellow }
    if ($skipped  -gt 0) { Write-Host "  skipped:  $skipped"  -ForegroundColor DarkGray }

    exit 0
}
catch {
    Write-Error "Install failed: $($_.Exception.Message)"
    exit 1
}
finally {
    if ($access) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit() } catch {}
        Release-Com $access
    }
    [System.GC]::Collect() | Out-Null
    [System.GC]::WaitForPendingFinalizers()
    if ($StageDir -and (Test-Path -LiteralPath $StageDir)) {
        Remove-Item $StageDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
