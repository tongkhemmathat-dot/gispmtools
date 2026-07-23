<#
    PMtools - interactive launcher.

    Shown when Run-PM.cmd is double-clicked. Lets the operator choose how long
    to sample CPU and memory before the assessment runs, without having to
    remember command-line switches.

    Deliberately English/ASCII, unlike the report itself. The Windows console
    renders Thai as boxes or question marks on most servers - the default raster
    font has no Thai glyphs and the console code page is rarely 874 - so a Thai
    menu would be unreadable exactly where it matters. Every other console
    message in PMtools is English for the same reason; the HTML report is where
    the Thai lives.

    Start-PMCheck.ps1 and Start-PMMonitor.ps1 remain callable directly for
    scripted or scheduled use; this file only wraps them.
#>
[CmdletBinding()]
param(
    # Left empty here on purpose - see the note below the param block.
    [string]$OutputRoot,
    [string]$ConfigDir
)

$ErrorActionPreference = 'Stop'

# $PSScriptRoot is empty while param() defaults are evaluated when a
# [CmdletBinding()] script runs under `powershell.exe -File` on PS 5.1 - which
# is precisely how Run-PM.cmd launches this file. Resolve it in the body.
# See the fuller note in Start-PMCheck.ps1.
$PMRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($PMRoot)) { $PMRoot = $PSScriptRoot }

if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Join-Path $PMRoot 'Output' }
if ([string]::IsNullOrWhiteSpace($ConfigDir))  { $ConfigDir  = Join-Path $PMRoot 'Config' }

. (Join-Path $PMRoot 'Lib\Core.ps1')
. (Join-Path $PMRoot 'Lib\ArcGIS.ps1')
Initialize-PMCore -ConfigDir $ConfigDir

$defaultMinutes = [int](Get-PMSetting -Path 'Monitor.DefaultMinutes'  -Default 30)
$interval       = [int](Get-PMSetting -Path 'Monitor.IntervalSeconds' -Default 10)
$maxAgeHours    = [double](Get-PMSetting -Path 'Monitor.MaxDataAgeHours' -Default 24)
$toolVersion    = [string](Get-PMSetting -Path 'Report.ToolVersion' -Default '1.2.0')

function Write-PMRule { Write-Host ('=' * 62) -ForegroundColor DarkGray }

function Get-PMFinishText {
    param([int]$Minutes)
    return ("{0} min, finishes about {1}" -f $Minutes, (Get-Date).AddMinutes($Minutes).ToString('HH:mm'))
}

# Report on sample data already on disk. The TREND check will happily use a file
# collected earlier today, so "run now" is often the right answer and the
# operator should be able to see that before choosing.
function Show-PMExistingData {
    $perfDir = Join-Path $OutputRoot '_Perf'
    if (-not (Test-Path -LiteralPath $perfDir)) { return }

    $csv = Get-ChildItem -LiteralPath $perfDir -Filter '*-perf-*.csv' -ErrorAction SilentlyContinue |
           Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $csv) { return }

    $ageHours = ((Get-Date) - $csv.LastWriteTime).TotalHours
    $count    = 0
    try { $count = @(Import-Csv -LiteralPath $csv.FullName -Encoding UTF8).Count } catch { $count = 0 }

    Write-Host '  Existing sample data: ' -NoNewline
    if ($ageHours -le $maxAgeHours) {
        Write-Host ("{0} samples, collected {1}" -f $count, $csv.LastWriteTime.ToString('HH:mm on yyyy-MM-dd')) -ForegroundColor Green
        Write-Host '                        option 1 will still chart it.' -ForegroundColor DarkGray
    }
    else {
        Write-Host ("{0:N1} h old - too old to chart (limit {1} h)" -f $ageHours, $maxAgeHours) -ForegroundColor Yellow
    }
    Write-Host ''
}

function Read-PMCustomMinutes {
    while ($true) {
        Write-Host ''
        $raw = Read-Host '  Sampling duration in minutes (1-1440, blank to cancel)'
        if ([string]::IsNullOrWhiteSpace($raw)) { return 0 }

        $value = 0
        if ([int]::TryParse($raw.Trim(), [ref]$value) -and $value -ge 1 -and $value -le 1440) { return $value }
        Write-Host '  Please enter a whole number between 1 and 1440.' -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------
# ArcGIS Server connection
#   English/ASCII like the rest of this menu - the Windows Server console
#   has no Thai glyphs. See the note at the top of this file.
# ---------------------------------------------------------------------

function Show-PMArcGISStatus {
    try { $conn = Get-PMArcGISConnection }
    catch {
        Write-Host '  Saved connection: ' -NoNewline
        Write-Host 'unreadable' -ForegroundColor Red
        Write-Host ('    ' + $_.Exception.Message) -ForegroundColor DarkGray
        return
    }

    Write-Host '  Saved connection: ' -NoNewline
    if ($null -eq $conn) {
        Write-Host 'none' -ForegroundColor Yellow
        Write-Host '    The ArcGIS checks stay skipped until one is set.' -ForegroundColor DarkGray
    }
    else {
        Write-Host $conn.Url -ForegroundColor Green
        Write-Host ("    user {0}, saved {1} by {2}" -f $conn.Username, $conn.SavedAt, $conn.SavedBy) -ForegroundColor DarkGray
    }
}

function Read-PMArcGISUrl {
    Write-Host ''
    Write-Host '  ArcGIS Server site URL' -ForegroundColor Cyan
    Write-Host '  Enter the site address WITHOUT the /admin part - examples:' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '    https://gis.example.go.th/server' -ForegroundColor Gray -NoNewline
    Write-Host '          through a Web Adaptor (most common)' -ForegroundColor DarkGray
    Write-Host '    https://gisserver.example.go.th:6443' -ForegroundColor Gray -NoNewline
    Write-Host '      straight to the server' -ForegroundColor DarkGray
    Write-Host '    https://localhost:6443' -ForegroundColor Gray -NoNewline
    Write-Host '                    when running on the server itself' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  A trailing /admin is accepted and removed for you.' -ForegroundColor DarkGray
    Write-Host ''

    $raw = Read-Host '  Site URL (blank to cancel)'
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }

    try { return Get-PMArcGISRoot -Url $raw }
    catch {
        Write-Host ('  ' + $_.Exception.Message) -ForegroundColor Yellow
        return $null
    }
}

function Set-PMArcGISConnectionInteractive {
    $url = Read-PMArcGISUrl
    if ($null -eq $url) { return }

    Write-Host ''
    Write-Host ('  Resolved to: {0}' -f $url) -ForegroundColor Green
    Write-Host ('  Admin API  : {0}/admin' -f $url) -ForegroundColor DarkGray

    Write-Host ''
    Write-Host '  Account' -ForegroundColor Cyan
    Write-Host '  Use an ArcGIS Server account that can READ the site. Examples:' -ForegroundColor DarkGray
    Write-Host '    siteadmin              built-in primary site administrator' -ForegroundColor Gray
    Write-Host '    pmreader               a dedicated read-only account (preferred)' -ForegroundColor Gray
    Write-Host '    DOMAIN\gis_monitor     when the site uses Windows accounts' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  PMtools only ever reads. A least-privilege account is enough and' -ForegroundColor DarkGray
    Write-Host '  is safer than the primary site administrator.' -ForegroundColor DarkGray
    Write-Host ''

    $user = Read-Host '  Username (blank to cancel)'
    if ([string]::IsNullOrWhiteSpace($user)) { return }

    Write-Host ''
    Write-Host '  The password is not shown as you type.' -ForegroundColor DarkGray
    $pass = Read-Host '  Password' -AsSecureString
    if ($null -eq $pass -or $pass.Length -eq 0) {
        Write-Host '  Cancelled - no password entered.' -ForegroundColor Yellow
        return
    }

    Write-Host ''
    Write-Host '  Testing the connection...' -ForegroundColor DarkGray
    $test = Test-PMArcGISConnection -Url $url -Username $user.Trim() -Password $pass

    Write-Host ''
    if (-not $test.Success) {
        Write-Host '  Connection FAILED' -ForegroundColor Red
        Write-Host ('    ' + $test.Message) -ForegroundColor Yellow
        Write-Host ''
        $anyway = Read-Host '  Save it anyway? (y/N)'
        if ($anyway.Trim().ToUpper() -ne 'Y') {
            Write-Host '  Not saved.' -ForegroundColor Yellow
            return
        }
    }
    else {
        Write-Host '  Connection OK' -ForegroundColor Green
        if ($test.Version)      { Write-Host ('    ArcGIS Server {0}' -f $test.Version) -ForegroundColor DarkGray }
        if ($test.MachineCount) { Write-Host ('    {0} machine(s) in the site' -f $test.MachineCount) -ForegroundColor DarkGray }
        if ($test.Message -ne 'Connected.') { Write-Host ('    ' + $test.Message) -ForegroundColor Yellow }
    }

    $path = Save-PMArcGISConnection -Url $url -Username $user.Trim() -Password $pass

    Write-Host ''
    Write-Host ('  Saved to {0}' -f $path) -ForegroundColor Green
    Write-Host '  The password is encrypted with Windows DPAPI: it can only be read' -ForegroundColor DarkGray
    Write-Host '  back by this Windows account on this machine.' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '  The ArcGIS checks are disabled by default. Run them with:' -ForegroundColor DarkGray
    Write-Host '    .\Start-PMCheck.ps1 -Only AGS' -ForegroundColor Gray
    Write-Host '  or remove "AGS" from Checks.Disabled in Config\settings.json to' -ForegroundColor DarkGray
    Write-Host '  include them in every run.' -ForegroundColor DarkGray
}

function Test-PMArcGISSavedConnection {
    try { $conn = Get-PMArcGISConnection }
    catch {
        Write-Host ''
        Write-Host '  Could not read the saved connection:' -ForegroundColor Red
        Write-Host ('    ' + $_.Exception.Message) -ForegroundColor Yellow
        return
    }

    if ($null -eq $conn) {
        Write-Host ''
        Write-Host '  Nothing saved yet - choose 1 first.' -ForegroundColor Yellow
        return
    }

    Write-Host ''
    Write-Host ('  Testing {0} as {1}...' -f $conn.Url, $conn.Username) -ForegroundColor DarkGray
    $test = Test-PMArcGISConnection -Url $conn.Url -Username $conn.Username -Password $conn.Password

    Write-Host ''
    if ($test.Success) {
        Write-Host '  Connection OK' -ForegroundColor Green
        if ($test.Version)      { Write-Host ('    ArcGIS Server {0}' -f $test.Version) -ForegroundColor DarkGray }
        if ($test.MachineCount) { Write-Host ('    {0} machine(s) in the site' -f $test.MachineCount) -ForegroundColor DarkGray }
        if ($test.Message -ne 'Connected.') { Write-Host ('    ' + $test.Message) -ForegroundColor Yellow }
    }
    else {
        Write-Host '  Connection FAILED' -ForegroundColor Red
        Write-Host ('    ' + $test.Message) -ForegroundColor Yellow
    }
}

function Show-PMArcGISMenu {
    while ($true) {
        Clear-Host
        Write-PMRule
        Write-Host '  ArcGIS Server connection' -ForegroundColor Cyan
        Write-PMRule
        Write-Host ''
        Show-PMArcGISStatus
        Write-Host ''
        Write-Host '   [1]  Set connection (URL, username, password)'
        Write-Host '   [2]  Test the saved connection'
        Write-Host '   [3]  Remove the saved connection'
        Write-Host ''
        Write-Host '   [B]  Back to the main menu'
        Write-Host ''

        $c = Read-Host '  Choice'
        switch ($c.Trim().ToUpper()) {
            '1' { Set-PMArcGISConnectionInteractive; Write-Host ''; Read-Host '  Press Enter to continue' | Out-Null }
            '2' { Test-PMArcGISSavedConnection;      Write-Host ''; Read-Host '  Press Enter to continue' | Out-Null }
            '3' {
                Write-Host ''
                if (Clear-PMArcGISConnection) {
                    Write-Host '  Removed.' -ForegroundColor Green
                }
                else {
                    Write-Host '  Nothing was saved.' -ForegroundColor Yellow
                }
                Write-Host ''
                Read-Host '  Press Enter to continue' | Out-Null
            }
            'B' { return }
            ''  { return }
            default {
                Write-Host ''
                Write-Host '  Not a valid choice.' -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
        }
    }
}

# --- what to check -----------------------------------------------------------
# A server is either a plain Windows Server with nothing to do with ArcGIS, or
# a GIS server where the ArcGIS checks are the point - never a useful mix of
# both in one run. This screen makes that an explicit, exclusive choice
# up front, rather than something buried in a command-line switch: everything
# past this point runs ONE of the two groups, never both. Start-PMCheck.ps1's
# own -Group parameter enforces the same split for anyone calling it directly.
#
# Nested in an outer $ready loop rather than recursing back into this script:
# choosing [B] on the sampling screen below needs to redraw the mode screen,
# and re-invoking the whole file for that would pile up nested script frames
# for no reason a loop does not already handle.
#
# Same break/continue-binds-to-switch trap noted on the inner loop below
# applies here too - see that comment for the fuller explanation.
$mode      = $null
$minutes   = 0
$runReport = $true
$ready     = $false

while (-not $ready) {

    $mode = $null
    while (-not $mode) {

        Clear-Host
        Write-PMRule
        Write-Host ("  PMtools {0} - Preventive Maintenance" -f $toolVersion) -ForegroundColor Cyan
        Write-Host ("  Server: {0}      {1}" -f $env:COMPUTERNAME, (Get-Date).ToString('yyyy-MM-dd HH:mm'))
        Write-PMRule
        Write-Host ''

        if (-not (Test-PMIsAdministrator)) {
            Write-Host '  NOTE: not running as Administrator - some checks will be' -ForegroundColor Yellow
            # Deliberately does not name a launcher: this same menu ships both
            # inside the PMtools folder (Run-PM.cmd) and inside the single-file
            # build (PMtools-<version>.cmd), and both elevate on their own.
            Write-Host '        incomplete. Close this and re-run the launcher to elevate.' -ForegroundColor Yellow
            Write-Host ''
        }

        # Read fresh each pass so the hint updates immediately after the
        # operator sets or removes a connection in the submenu.
        $agsHint = '(not configured)'
        try {
            $agsConn = Get-PMArcGISConnection
            if ($null -ne $agsConn) { $agsHint = '(' + $agsConn.Url + ')' }
        }
        catch { $agsHint = '(saved, but unreadable here)' }

        Write-Host '  What do you want to check?'
        Write-Host ''
        Write-Host '   [1]  Server - the regular maintenance checks'
        Write-Host '   [2]  ArcGIS Server - site, services, usage reports  ' -NoNewline
        Write-Host $agsHint -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '   [A]  ArcGIS Server connection...'
        Write-Host ''
        Write-Host '   [Q]  Quit'
        Write-Host ''

        $modeChoice = Read-Host '  Choice [1]'
        if ([string]::IsNullOrWhiteSpace($modeChoice)) { $modeChoice = '1' }

        switch ($modeChoice.Trim().ToUpper()) {
            '1' { $mode = 'Server' }
            '2' {
                $conn = $null
                try { $conn = Get-PMArcGISConnection } catch { $conn = $null }
                if ($null -eq $conn) {
                    Write-Host ''
                    Write-Host '  No ArcGIS Server connection is configured yet - set one up first.' -ForegroundColor Yellow
                    Write-Host ''
                    Read-Host '  Press Enter to continue' | Out-Null
                    Show-PMArcGISMenu
                }
                else {
                    $mode = 'ArcGIS'
                }
            }
            'A' { Show-PMArcGISMenu }
            'Q' { Write-Host ''; Write-Host '  Cancelled.'; Write-Host ''; exit 0 }
            default {
                Write-Host ''
                Write-Host '  Not a valid choice.' -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
        }
    }

    if ($mode -eq 'ArcGIS') { $ready = $true; continue }

    # --- Server mode: how long to sample first -------------------------------
    # $chosen drives the loop rather than break/continue inside the switch.
    #
    # PowerShell binds break and continue to the SWITCH, not to an enclosing
    # while: a `continue` in a case exits the switch and falls straight into
    # whatever follows it. An earlier version of this menu ended `while` with
    # a bare `break` after the switch and used `continue` to mean "show the
    # menu again", so pressing an invalid key - or cancelling out of the
    # custom duration prompt - silently started a full assessment instead of
    # redrawing the menu. Verified against the live behaviour, not assumed.
    $backToModeScreen = $false
    $chosen = $false

    while (-not $chosen) {

        Clear-Host
        Write-PMRule
        Write-Host ("  PMtools {0} - Preventive Maintenance" -f $toolVersion) -ForegroundColor Cyan
        Write-Host ("  Server: {0}      {1}" -f $env:COMPUTERNAME, (Get-Date).ToString('yyyy-MM-dd HH:mm'))
        Write-PMRule
        Write-Host ''

        Show-PMExistingData

        Write-Host '  Collect CPU / memory samples before the assessment?'
        Write-Host ''
        Write-Host '   [1]  No sampling - assess now                 ' -NoNewline
        Write-Host '(about 10 s)' -ForegroundColor DarkGray
        Write-Host ('   [2]  Sample 15 minutes, then assess           ({0})' -f (Get-PMFinishText 15))
        Write-Host ('   [3]  Sample {0} minutes, then assess           ({1})' -f $defaultMinutes, (Get-PMFinishText $defaultMinutes))
        Write-Host ('   [4]  Sample 60 minutes, then assess           ({0})' -f (Get-PMFinishText 60))
        Write-Host '   [5]  Sample for a custom duration...'
        Write-Host '   [6]  Sample only - do not build a report'
        Write-Host ''
        Write-Host '   [B]  Back'
        Write-Host '   [Q]  Quit'
        Write-Host ''

        $choice = Read-Host '  Choice [1]'
        if ([string]::IsNullOrWhiteSpace($choice)) { $choice = '1' }

        # Every case either sets $chosen (proceed) or leaves it false (redraw).
        switch ($choice.Trim().ToUpper()) {
            '1' { $minutes = 0;               $runReport = $true;  $chosen = $true }
            '2' { $minutes = 15;              $runReport = $true;  $chosen = $true }
            '3' { $minutes = $defaultMinutes; $runReport = $true;  $chosen = $true }
            '4' { $minutes = 60;              $runReport = $true;  $chosen = $true }
            '5' {
                $custom = Read-PMCustomMinutes
                if ($custom -gt 0) { $minutes = $custom; $runReport = $true; $chosen = $true }
            }
            '6' {
                $custom = Read-PMCustomMinutes
                if ($custom -gt 0) { $minutes = $custom; $runReport = $false; $chosen = $true }
            }
            'B' { $backToModeScreen = $true; $chosen = $true }
            'Q' { Write-Host ''; Write-Host '  Cancelled.'; Write-Host ''; exit 0 }
            default {
                Write-Host ''
                Write-Host '  Not a valid choice.' -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
        }
    }

    if (-not $backToModeScreen) { $ready = $true }
}

# --- act on it --------------------------------------------------------------
Write-Host ''

if ($mode -eq 'Server') {
    if ($minutes -gt 0) {
        & (Join-Path $PMRoot 'Start-PMMonitor.ps1') `
            -Minutes $minutes -IntervalSeconds $interval -OutputRoot $OutputRoot -ConfigDir $ConfigDir

        if (-not $runReport) {
            Write-Host 'Sampling finished. Run Run-PM.cmd again and choose 1 to build the report.'
            Write-Host ''
            exit 0
        }
    }

    & (Join-Path $PMRoot 'Start-PMCheck.ps1') -Group Server -OutputRoot $OutputRoot -ConfigDir $ConfigDir -OpenReport
    exit $LASTEXITCODE
}

# ArcGIS mode: no sampling question, nothing to trend - just run the three
# ArcGIS checks against the connection confirmed to exist above.
& (Join-Path $PMRoot 'Start-PMCheck.ps1') -Group ArcGIS -OutputRoot $OutputRoot -ConfigDir $ConfigDir -OpenReport
exit $LASTEXITCODE
