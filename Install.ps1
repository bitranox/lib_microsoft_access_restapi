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

function Wait-ForKey {
    # When launched by double-click the window would close immediately
    # otherwise. Fall back to Read-Host when no interactive RawUI is
    # available (ISE, piped input, non-console host).
    Write-Host ""
    try {
        if ($Host.UI.RawUI -and $Host.Name -eq 'ConsoleHost') {
            Write-Host "Press any key to exit..." -ForegroundColor Cyan
            [void]$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        } else {
            Read-Host "Press Enter to exit"
        }
    } catch {
        Read-Host "Press Enter to exit"
    }
}

function Exit-WithKey([int]$code) {
    Wait-ForKey
    exit $code
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

# When invoked via `powershell -File ... -Database "C:\path with spaces\f.accdb"`
# from cmd.exe, the surrounding quotes are sometimes passed *into* the
# parameter value (so $Database literally starts and ends with "). Strip
# stray surrounding quotes and whitespace so Test-Path actually sees the path.
$RawDatabase = $Database
$Database = $Database.Trim()
if ($Database.Length -ge 2) {
    $first = $Database[0]; $last = $Database[$Database.Length - 1]
    if (($first -eq '"' -and $last -eq '"') -or ($first -eq "'" -and $last -eq "'")) {
        $Database = $Database.Substring(1, $Database.Length - 2).Trim()
    }
}

if (-not (Test-Path -LiteralPath $Database)) {
    Write-Host ""
    Write-Host "ERROR: The specified database file was not found:" -ForegroundColor Red
    Write-Host "  [$Database]" -ForegroundColor Red
    if ($RawDatabase -ne $Database) {
        Write-Host "  (raw argument received: [$RawDatabase])" -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "Please check the path and run the script again." -ForegroundColor Yellow
    if ($Database -match '\s' -or $Database -match '[&()]') {
        Write-Host ""
        Write-Host "Tip: paths containing spaces or special characters must be quoted." -ForegroundColor Yellow
        Write-Host "From cmd.exe, prefer the -Command form over -File:" -ForegroundColor Yellow
        Write-Host '  powershell -ExecutionPolicy Bypass -Command "& .\Install.ps1 -Database ''C:\path with spaces\app.accdb''"' -ForegroundColor Yellow
    }
    Exit-WithKey 2
}

$Database = (Resolve-Path -LiteralPath $Database).Path

foreach ($m in $Modules) {
    $path = Join-Path $ScriptDir $m
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Host ""
        Write-Host "ERROR: Missing module file:" -ForegroundColor Red
        Write-Host "  $path" -ForegroundColor Red
        Exit-WithKey 2
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
    Write-Host ""
    Write-Host "ERROR: Could not start Access via COM." -ForegroundColor Red
    Write-Host "Is Microsoft Access installed?" -ForegroundColor Red
    Exit-WithKey 3
}

$ExitCode = 0
$StageDir = $null

try {
    $access.Visible = [bool]$Visible
    try { $access.AutomationSecurity = 1 } catch {}  # msoAutomationSecurityLow
    try { $access.DoCmd.SetWarnings($false) } catch {}

    Write-Host "Opening database..." -ForegroundColor Cyan
    try {
        $access.OpenCurrentDatabase($Database)
    } catch {
        Write-Host ""
        Write-Host "ERROR: Failed to open database:" -ForegroundColor Red
        Write-Host "  $Database" -ForegroundColor Red
        Write-Host ""
        Write-Host "Common causes:" -ForegroundColor Yellow
        Write-Host "  - The file is already open in another Access window. Close it and retry."
        Write-Host "  - The path or file is wrong / corrupt."
        Write-Host "  - The file is read-only or locked by another process."
        Write-Host ""
        Write-Host "Underlying error: $($_.Exception.Message)" -ForegroundColor DarkGray
        $ExitCode = 4
        return
    }

    try {
        $vbe = $access.VBE
        $vbproj = $vbe.VBProjects.Item(1)
    } catch {
        Write-Host ""
        Write-Host "ERROR: Cannot reach the VBA project model." -ForegroundColor Red
        Write-Host ""
        Write-Host "In Access, open File -> Options -> Trust Center -> Trust Center Settings ->" -ForegroundColor Yellow
        Write-Host "Macro Settings, and tick 'Trust access to the VBA project object model'." -ForegroundColor Yellow
        Write-Host "Restart Access and run this script again." -ForegroundColor Yellow
        $ExitCode = 5
        return
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
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host " Installation completed successfully." -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "  Database: $Database" -ForegroundColor Green
    Write-Host "  imported: $imported" -ForegroundColor Green
    if ($replaced -gt 0) { Write-Host "  replaced: $replaced" -ForegroundColor Yellow }
    if ($skipped  -gt 0) { Write-Host "  skipped:  $skipped"  -ForegroundColor DarkGray }
    $ExitCode = 0
}
catch {
    Write-Host ""
    Write-Host "ERROR: Install failed." -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    $ExitCode = 1
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

Wait-ForKey
exit $ExitCode
