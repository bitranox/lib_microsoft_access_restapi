<#
.SYNOPSIS
    Self-contained runner for lib_access_restapi tests.

.DESCRIPTION
    Drives Microsoft Access via COM automation to:
      1. Create a throwaway .accdb next to this script
      2. Import every .bas / .cls module from the script directory
      3. Add the MSXML 6.0 type-library reference (belt-and-suspenders;
         the library itself is fully late-bound via CreateObject)
      4. Run modRestTest.Test_All, capturing PASS/FAIL lines to a file
      5. Print captured output and exit with the FAIL count as exit code

    Run from the directory containing the modules:

        powershell -ExecutionPolicy Bypass -File .\Run-Tests.ps1

    Or pipe straight into the host's PS prompt:

        .\Run-Tests.ps1 -JsonOnly        # skip HTTP tests (no network needed)

.PARAMETER JsonOnly
    Only run the offline JSON unit tests. Skips the httpbin.org calls.
    Useful for smoke-testing on a box without internet.

.PARAMETER KeepDatabase
    Don't delete the temporary .accdb on exit. Path is printed at the end.
    Handy for poking at the project in the VBE after a failure.

.NOTES
    Prereqs:
      - Microsoft Access (64-bit recommended, matches the library target)
      - "Trust access to the VBA project object model" must be enabled in
        Access Trust Center -> Macro Settings. Without it, VBE automation
        raises an "Access is denied" / "Programmatic access" error.
      - Internet access to httpbin.org for the HTTP tests (skip with
        -JsonOnly).
#>

[CmdletBinding()]
param(
    [switch] $JsonOnly,
    [switch] $KeepDatabase,
    [switch] $Visible
)

$ErrorActionPreference = 'Stop'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Files to import. Order matters only when later modules reference earlier
# ones at compile time -- VBA does forward resolution so the order below
# is just "dependencies first" for readability.
$Modules = @(
    'modJson.bas',
    'IHttpCallback.cls',
    'clsHttpResponse.cls',
    'clsAsyncCall.cls',
    'clsHttpRequest.cls',
    'modRestPump.bas',
    'modRest.bas',
    'clsResponseCollector.cls',
    'clsLateBoundReceiver.cls',
    'modRestTest.bas'
)

# MSXML 6.0 type library GUID. Stable across Windows versions; ships in OS.
$MsxmlGuid  = '{F5078F18-C551-11D3-89B9-0000F81FE221}'
$MsxmlMajor = 6
$MsxmlMinor = 0

# Put the temp DB and log next to the script (in the repo dir), not in
# %TEMP%. Easier to inspect after a failed run, and avoids redirected /
# AV-scanned %TEMP% locations that can cause odd "Save As" prompts.
$stamp      = [Guid]::NewGuid().ToString('N')
$DbPath     = Join-Path $ScriptDir "restapi_test_$stamp.accdb"
$ResultPath = Join-Path $ScriptDir "restapi_test_$stamp.log"

function Release-Com([object]$obj) {
    if ($obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
    }
}

# Return a path to a CRLF-normalized copy of $srcPath. The VBE rejects
# LF-only .bas/.cls files with "Erwartet Anweisungsende" because it
# parses the VERSION/BEGIN/END/Attribute header expecting CRLF. We copy
# into the script directory's _stage subfolder which is cleaned up at
# the end of the run.
function Convert-ToCrlf([string]$srcPath, [string]$stageDir) {
    if (-not (Test-Path -LiteralPath $stageDir)) {
        New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
    }
    $dstPath = Join-Path $stageDir (Split-Path -Leaf $srcPath)
    # Read raw bytes so we don't accidentally double-encode UTF-8 BOM etc.
    $raw = [System.IO.File]::ReadAllText($srcPath)
    # Normalize: collapse any existing CRLF to LF first, then re-emit CRLF.
    $normalized = ($raw -replace "`r`n", "`n") -replace "`n", "`r`n"
    [System.IO.File]::WriteAllText($dstPath, $normalized, (New-Object System.Text.UTF8Encoding $false))
    return $dstPath
}

# Sanity-check that every required file is present
foreach ($m in $Modules) {
    $path = Join-Path $ScriptDir $m
    if (-not (Test-Path $path)) {
        Write-Error "Missing module file: $path"
        exit 2
    }
}

Write-Host "Starting Access..." -ForegroundColor Cyan
$access = $null
try {
    $access = New-Object -ComObject Access.Application
} catch {
    Write-Error "Could not start Access via COM. Is Microsoft Access installed?"
    exit 3
}

try {
    $access.Visible = [bool]$Visible
    # Suppress macro / "unsafe content" prompts that would otherwise hang
    # an invisible Access waiting for a click.
    try { $access.AutomationSecurity = 1 } catch {}  # msoAutomationSecurityLow
    try { $access.DoCmd.SetWarnings($false) } catch {}

    Write-Host "Creating temp database: $DbPath" -ForegroundColor Cyan
    $access.NewCurrentDatabase($DbPath)

    # Reach the VBA project. Requires Trust Center -> "Trust access to
    # the VBA project object model" to be enabled.
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
        exit 4
    }

    Write-Host "Adding MSXML 6.0 reference..." -ForegroundColor Cyan
    try {
        $vbproj.References.AddFromGuid($MsxmlGuid, $MsxmlMajor, $MsxmlMinor) | Out-Null
    } catch {
        # Already present? Continue. Anything else: surface it.
        if ($_.Exception.Message -notmatch 'already') { throw }
    }

    Write-Host "Importing modules..." -ForegroundColor Cyan
    $StageDir = Join-Path $ScriptDir "_stage_$stamp"
    foreach ($m in $Modules) {
        $src = Join-Path $ScriptDir $m
        Write-Host "  + $m"
        $crlfPath = Convert-ToCrlf $src $StageDir
        $vbproj.VBComponents.Import($crlfPath) | Out-Null
    }

    # Don't call DoCmd.RunCommand(126) here: the constant ID isn't
    # consistent across Access versions and on some setups it triggers
    # the VBE's "Export Module As Text" dialog instead of a silent
    # compile-and-save. The first Application.Run call below will force
    # the project to compile; if there's a compile error it surfaces
    # there as a clean exception.

    Write-Host "Wiring log capture: $ResultPath" -ForegroundColor Cyan
    try {
        $access.Run('SetLogFile', $ResultPath)
    } catch {
        Write-Error @"
SetLogFile failed. Most likely a VBA compile error or a missing reference.

If Run-Tests.ps1 hangs at this step instead of erroring, re-run with -Visible
to see what dialog Access is showing.

Underlying error: $($_.Exception.Message)
"@
        throw
    }

    if ($JsonOnly) {
        Write-Host "Running JSON tests only..." -ForegroundColor Cyan
        $access.Run('Test_Json_All')
    } else {
        Write-Host "Running full test suite (this hits httpbin.org)..." -ForegroundColor Cyan
        $access.Run('Test_All')
    }

    $failCount = [int] $access.Run('FailCount')

    Write-Host ""
    Write-Host "----- Test output -----" -ForegroundColor Yellow
    if (Test-Path $ResultPath) {
        Get-Content $ResultPath | ForEach-Object {
            if ($_ -match '^\s*FAIL:') {
                Write-Host $_ -ForegroundColor Red
            } elseif ($_ -match '^\s*PASS:') {
                Write-Host $_ -ForegroundColor Green
            } else {
                Write-Host $_
            }
        }
    } else {
        Write-Warning "No log file was produced."
    }
    Write-Host "-----------------------" -ForegroundColor Yellow

    if ($failCount -eq 0) {
        Write-Host "ALL GREEN" -ForegroundColor Green
    } else {
        Write-Host "$failCount FAILURE(S)" -ForegroundColor Red
    }

    exit $failCount
}
catch {
    Write-Host ""
    Write-Error "Run failed: $($_.Exception.Message)"
    if (Test-Path $ResultPath) {
        Write-Host "--- Partial log ---" -ForegroundColor Yellow
        Get-Content $ResultPath | ForEach-Object { Write-Host $_ }
    }
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

    $runOk = ($LASTEXITCODE -eq 0)

    if ($KeepDatabase -or -not $runOk) {
        # Always preserve the DB on failure so the user can open it in Access
        # and run Debug -> Compile to see the actual compile error.
        Write-Host ""
        Write-Host "Kept database at: $DbPath" -ForegroundColor DarkGray
        Write-Host "Log at:           $ResultPath" -ForegroundColor DarkGray
        if (-not $runOk) {
            Write-Host ""
            Write-Host "To see the underlying VBA error:" -ForegroundColor Yellow
            Write-Host "  1. Open the database above in Access." -ForegroundColor Yellow
            Write-Host "  2. Press Alt+F11 to open the VBE." -ForegroundColor Yellow
            Write-Host "  3. Debug -> Compile (the first highlighted line is the cause)." -ForegroundColor Yellow
        }
    } else {
        Remove-Item $DbPath     -Force -ErrorAction SilentlyContinue
        Remove-Item $ResultPath -Force -ErrorAction SilentlyContinue
    }
    # Always clean up the stage dir
    if ($StageDir -and (Test-Path -LiteralPath $StageDir)) {
        Remove-Item $StageDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}
