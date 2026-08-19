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
$toolVersion    = [string](Get-PMSetting -Path 'Report.ToolVersion' -Default '1.7.13')

function Write-PMRule { Write-Host ('=' * 62) -ForegroundColor DarkGray }

# Shared rendering for every "[N]  Description" menu line, so the key that
# gets pressed is visually distinct from its description everywhere instead
# of blending into the same plain-color text - and so the bare-Enter default
# is marked right on the list, not only in the "Choice [1]" prompt below it.
function Write-PMMenuOption {
    param([string]$Key, [string]$Text, [switch]$Default, [switch]$NoNewline)
    Write-Host ('   [{0}]  ' -f $Key) -ForegroundColor White -NoNewline
    if ($Default) {
        Write-Host $Text -NoNewline
        if (-not $NoNewline) { Write-Host '  (default)' -ForegroundColor DarkGray }
    }
    elseif ($NoNewline) { Write-Host $Text -NoNewline }
    else { Write-Host $Text }
}

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

# Same shape as Read-PMCustomMinutes, seconds instead of minutes, for the NAS
# ad-hoc test interval. Capped at 600 (10 min) - anything looser than that is
# really a "sample and chart it" job like Start-PMMonitor.ps1, not a live
# console loop someone is watching in real time.
function Read-PMCustomSeconds {
    while ($true) {
        Write-Host ''
        $raw = Read-Host '  Interval in seconds (1-600, blank to cancel)'
        if ([string]::IsNullOrWhiteSpace($raw)) { return 0 }

        $value = 0
        if ([int]::TryParse($raw.Trim(), [ref]$value) -and $value -ge 1 -and $value -le 600) { return $value }
        Write-Host '  Please enter a whole number between 1 and 600.' -ForegroundColor Yellow
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
        if ($conn.AuthMode -eq 'Portal') {
            Write-Host ("    auth via Portal: {0}" -f $conn.PortalUrl) -ForegroundColor DarkGray
        }
        Write-Host ("    user {0}, saved {1} by {2}" -f $conn.Username, $conn.SavedAt, $conn.SavedBy) -ForegroundColor DarkGray
    }
}

function Show-PMArcGISUrlHelp {
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
}

function Show-PMArcGISPortalUrlHelp {
    Write-Host '  Portal for ArcGIS URL' -ForegroundColor Cyan
    Write-Host '  Enter the Portal this server is federated with - examples:' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '    https://portal.example.go.th/portal' -ForegroundColor Gray -NoNewline
    Write-Host '   Web Adaptor in front of Portal (most common)' -ForegroundColor DarkGray
    Write-Host '    https://localhost/portal' -ForegroundColor Gray -NoNewline
    Write-Host '                when running on the Portal machine itself' -ForegroundColor DarkGray
    Write-Host ''
}

# Every step below redraws through this - Back always lands on a clean
# screen instead of printing under whatever was already there, same rhythm
# as the mode/sampling screens further down this file, just applied per step.
function Show-PMArcGISWizardHeader {
    param([int]$Step, [int]$TotalSteps)
    Clear-Host
    Write-PMRule
    Write-Host ('  ArcGIS Server connection - step {0} of {1}' -f $Step, $TotalSteps) -ForegroundColor Cyan
    Write-PMRule
    Write-Host ''
}

# Recaps whatever has already been decided, since each step now only shows
# its own prompt - without this the operator would lose sight of the site
# URL (and Portal URL, and account) they already entered on earlier steps.
function Show-PMArcGISWizardContext {
    param([string]$Url, [string]$AuthMode, [string]$PortalUrl, [string]$User)
    if ($Url)                                          { Write-Host ('  Site: {0}' -f $Url) -ForegroundColor DarkGray }
    if ($AuthMode -eq 'Portal' -and $PortalUrl)         { Write-Host ('  Portal: {0}' -f $PortalUrl) -ForegroundColor DarkGray }
    if ($User)                                          { Write-Host ('  Account: {0}' -f $User) -ForegroundColor DarkGray }
    Write-Host ''
}

function Set-PMArcGISConnectionInteractive {
    $url       = $null
    $authMode  = 'Server'
    $portalUrl = ''
    $user      = $null
    $pass      = $null

    # 1=URL 2=Federation 3=Portal URL (only when federated) 4=Username
    # 5=Password 6=done (falls out of the loop into the Test+Save below).
    #
    # "B"/"Q" are checked BEFORE anything is handed to a URL parser at every
    # step, so typing the single letter "b" can never be mistaken for a
    # (nonsensical but syntactically valid - Get-PMArcGISRoot prepends
    # https:// to anything scheme-less) URL like "https://b".
    #
    # Every case below only ever sets $step (or returns) and lets the
    # switch end normally - it never relies on `continue` to "restart the
    # loop". `continue` inside a switch binds to the switch, not to this
    # enclosing while, so it would just fall through past the switch
    # instead of redrawing - the exact break/continue-binds-to-switch trap
    # already documented above the sampling-duration loop further down
    # this file, hit once before and worth not repeating here.
    $step = 1
    while ($step -ge 1 -and $step -le 5) {
        switch ($step) {
            1 {
                Show-PMArcGISWizardHeader -Step 1 -TotalSteps 5
                Show-PMArcGISUrlHelp
                $raw = Read-Host '  Site URL (blank to cancel)'
                if ([string]::IsNullOrWhiteSpace($raw)) { return }

                try {
                    $url = Get-PMArcGISRoot -Url $raw
                    Write-Host ''
                    Write-Host ('  Resolved to: {0}' -f $url) -ForegroundColor Green
                    Write-Host ('  Admin API  : {0}/admin' -f $url) -ForegroundColor DarkGray
                    Write-Host ''
                    Read-Host '  Press Enter to continue' | Out-Null
                    $step = 2
                }
                catch {
                    Write-Host ('  ' + $_.Exception.Message) -ForegroundColor Yellow
                    Write-Host ''
                    Read-Host '  Press Enter to try again' | Out-Null
                }
            }
            2 {
                Show-PMArcGISWizardHeader -Step 2 -TotalSteps 5
                Show-PMArcGISWizardContext -Url $url
                Write-Host '  Federation' -ForegroundColor Cyan
                Write-Host '  Is this ArcGIS Server federated with a Portal for ArcGIS?' -ForegroundColor DarkGray
                Write-Host '  Answer yes if you sign in with a Portal named user rather than a' -ForegroundColor DarkGray
                Write-Host '  server-only account such as siteadmin.' -ForegroundColor DarkGray
                $fedAnswer = (Read-Host '  Federated with Portal? (y/N, B to go back, Q to cancel)').Trim().ToUpper()

                if     ($fedAnswer -eq 'B') { $step = 1 }
                elseif ($fedAnswer -eq 'Q') { Write-Host ''; Write-Host '  Cancelled.' -ForegroundColor Yellow; return }
                elseif ($fedAnswer -eq 'Y') { $authMode = 'Portal'; $step = 3 }
                else                        { $authMode = 'Server'; $portalUrl = ''; $step = 4 }
            }
            3 {
                Show-PMArcGISWizardHeader -Step 3 -TotalSteps 5
                Show-PMArcGISWizardContext -Url $url -AuthMode $authMode
                Show-PMArcGISPortalUrlHelp
                $raw = Read-Host '  Portal URL (blank or B to go back, Q to cancel)'
                $ans = $raw.Trim().ToUpper()

                if ([string]::IsNullOrWhiteSpace($raw) -or $ans -eq 'B') { $step = 2 }
                elseif ($ans -eq 'Q') { Write-Host ''; Write-Host '  Cancelled.' -ForegroundColor Yellow; return }
                else {
                    try {
                        $portalUrl = Get-PMArcGISPortalRoot -Url $raw
                        Write-Host ''
                        Write-Host ('  Resolved to: {0}' -f $portalUrl) -ForegroundColor Green
                        Write-Host ''
                        Read-Host '  Press Enter to continue' | Out-Null
                        $step = 4
                    }
                    catch {
                        Write-Host ('  ' + $_.Exception.Message) -ForegroundColor Yellow
                        Write-Host ''
                        Read-Host '  Press Enter to try again' | Out-Null
                    }
                }
            }
            4 {
                Show-PMArcGISWizardHeader -Step 4 -TotalSteps 5
                Show-PMArcGISWizardContext -Url $url -AuthMode $authMode -PortalUrl $portalUrl
                Write-Host '  Account' -ForegroundColor Cyan
                if ($authMode -eq 'Portal') {
                    Write-Host '  Sign in with a Portal member account that has administrative' -ForegroundColor DarkGray
                    Write-Host '  access to this federated server. A server-only account such as' -ForegroundColor DarkGray
                    Write-Host '  siteadmin will NOT work here - it does not exist in Portal.' -ForegroundColor DarkGray
                    Write-Host ''
                    Write-Host '  PMtools only ever reads. A least-privilege administrative role is' -ForegroundColor DarkGray
                    Write-Host '  enough and is safer than a full organization administrator.' -ForegroundColor DarkGray
                }
                else {
                    Write-Host '  Use an ArcGIS Server account that can READ the site. Examples:' -ForegroundColor DarkGray
                    Write-Host '    siteadmin              built-in primary site administrator' -ForegroundColor Gray
                    Write-Host '    pmreader               a dedicated read-only account (preferred)' -ForegroundColor Gray
                    Write-Host '    DOMAIN\gis_monitor     when the site uses Windows accounts' -ForegroundColor Gray
                    Write-Host ''
                    Write-Host '  PMtools only ever reads. A least-privilege account is enough and' -ForegroundColor DarkGray
                    Write-Host '  is safer than the primary site administrator.' -ForegroundColor DarkGray
                }
                Write-Host ''
                $raw = Read-Host '  Username (blank or B to go back, Q to cancel)'
                $ans = $raw.Trim().ToUpper()

                if ([string]::IsNullOrWhiteSpace($raw) -or $ans -eq 'B') {
                    $step = if ($authMode -eq 'Portal') { 3 } else { 2 }
                }
                elseif ($ans -eq 'Q') { Write-Host ''; Write-Host '  Cancelled.' -ForegroundColor Yellow; return }
                else { $user = $raw.Trim(); $step = 5 }
            }
            5 {
                Show-PMArcGISWizardHeader -Step 5 -TotalSteps 5
                Show-PMArcGISWizardContext -Url $url -AuthMode $authMode -PortalUrl $portalUrl -User $user
                Write-Host '  The password is not shown as you type.' -ForegroundColor DarkGray
                Write-Host '  Type B and Enter to go back, or Q and Enter to cancel.' -ForegroundColor DarkGray
                $secure = Read-Host '  Password' -AsSecureString

                $plain = $null
                if ($secure -and $secure.Length -gt 0) { $plain = ConvertFrom-PMSecureString -Secure $secure }

                if ([string]::IsNullOrWhiteSpace($plain)) {
                    $plain = $null
                    Write-Host ''
                    Write-Host '  No password entered - going back.' -ForegroundColor Yellow
                    Write-Host ''
                    Read-Host '  Press Enter to continue' | Out-Null
                    $step = 4
                }
                elseif ($plain.Trim().ToUpper() -eq 'B') {
                    $plain = $null
                    $step  = 4
                }
                elseif ($plain.Trim().ToUpper() -eq 'Q') {
                    $plain = $null
                    Write-Host ''
                    Write-Host '  Cancelled.' -ForegroundColor Yellow
                    return
                }
                else {
                    $pass  = $secure
                    $plain = $null
                    $step  = 6
                }
            }
        }
    }

    if ($step -ne 6) { return }

    Write-Host ''
    Write-Host '  Testing the connection...' -ForegroundColor DarkGray
    $test = Test-PMArcGISConnection -Url $url -Username $user -Password $pass `
                                    -AuthMode $authMode -PortalUrl $portalUrl

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

    $path = Save-PMArcGISConnection -Url $url -Username $user -Password $pass `
                                    -AuthMode $authMode -PortalUrl $portalUrl

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
    $test = Test-PMArcGISConnection -Url $conn.Url -Username $conn.Username -Password $conn.Password `
                                    -AuthMode $conn.AuthMode -PortalUrl $conn.PortalUrl

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
        Write-PMMenuOption -Key '1' -Text 'Set connection (URL, username, password)'
        Write-PMMenuOption -Key '2' -Text 'Test the saved connection'
        Write-PMMenuOption -Key '3' -Text 'Remove the saved connection'
        Write-Host ''
        Write-PMMenuOption -Key 'B' -Text 'Back to the main menu'
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

# Shared by both ArcGIS entries on the "what to check" screen below. Offers
# to set a connection up on the spot rather than just refusing, since a
# first-time operator landing on either entry has nowhere else obvious to go.
function Confirm-PMArcGISConnectionReady {
    $conn = $null
    try { $conn = Get-PMArcGISConnection } catch { $conn = $null }
    if ($null -ne $conn) { return $true }

    Write-Host ''
    Write-Host '  No ArcGIS Server connection is configured yet - set one up first.' -ForegroundColor Yellow
    Write-Host ''
    Read-Host '  Press Enter to continue' | Out-Null
    Show-PMArcGISMenu
    return $false
}

# ---------------------------------------------------------------------
# NAS / SMB connectivity testing
#   Two different jobs live under one menu entry: running the *configured*
#   SmbConnectivity.Targets through the normal SMBCONN check (same as any
#   other check - lands in the HTML report), and an ad-hoc console-only test
#   against a path typed in right now, optionally as a different user,
#   optionally repeated on an interval. The ad-hoc path never touches
#   Config\settings.json and never writes to the report - it exists for
#   "is this share reachable right now" during troubleshooting, not for
#   the assessment record.
# ---------------------------------------------------------------------

# Credentials are asked for only when a share path was actually given - the
# port test never authenticates, so a username/password would do nothing for
# a port-only test and asking for one anyway would just be confusing.
function Read-PMSmbCredential {
    Write-Host ''
    $user = Read-Host '  Username to authenticate as (blank = current account)'
    if ([string]::IsNullOrWhiteSpace($user)) { return $null }

    Write-Host '  Password is used once for this test only - never saved.' -ForegroundColor DarkGray
    $secure = Read-Host '  Password' -AsSecureString
    if (-not $secure -or $secure.Length -eq 0) { return $null }
    return New-Object System.Management.Automation.PSCredential($user.Trim(), $secure)
}

function Read-PMAdHocSmbTarget {
    Write-Host ''
    Write-Host '  Test a custom path' -ForegroundColor Cyan
    $server = Read-Host '  NAS server - hostname or IP (blank to cancel)'
    if ([string]::IsNullOrWhiteSpace($server)) { return $null }
    $server = $server.Trim()

    $portRaw = Read-Host '  Port [445]'
    $port = 445
    if (-not [string]::IsNullOrWhiteSpace($portRaw)) {
        $parsed = 0
        if ([int]::TryParse($portRaw.Trim(), [ref]$parsed) -and $parsed -gt 0 -and $parsed -le 65535) { $port = $parsed }
        else { Write-Host '  Invalid port - using 445.' -ForegroundColor Yellow }
    }

    $share = (Read-Host ('  Share path to test, e.g. \\{0}\share (blank = port only)' -f $server)).Trim()

    $cred = $null
    if (-not [string]::IsNullOrWhiteSpace($share)) { $cred = Read-PMSmbCredential }

    return [pscustomobject]@{ Server = $server; Port = $port; SharePath = $share; Credential = $cred }
}

# Returns @{ IntervalSeconds; DurationMinutes }, or $null for a single pass.
function Read-PMTestSchedule {
    Write-Host ''
    Write-Host '  Test once, or repeatedly over time?' -ForegroundColor Cyan
    Write-PMMenuOption -Key '1' -Text 'Once' -Default
    Write-PMMenuOption -Key '2' -Text 'Repeatedly...'
    Write-Host ''
    $c = Read-Host '  Choice [1]'
    if ([string]::IsNullOrWhiteSpace($c)) { $c = '1' }
    if ($c.Trim() -ne '2') { return $null }

    Write-Host ''
    Write-Host '  How often?' -ForegroundColor Cyan
    Write-PMMenuOption -Key '1' -Text 'Every 5 seconds'
    Write-PMMenuOption -Key '2' -Text 'Every 10 seconds' -Default
    Write-PMMenuOption -Key '3' -Text 'Every 30 seconds'
    Write-PMMenuOption -Key '4' -Text 'Custom interval...'
    Write-Host ''
    $ic = Read-Host '  Choice [2]'
    if ([string]::IsNullOrWhiteSpace($ic)) { $ic = '2' }
    $interval = switch ($ic.Trim()) {
        '1' { 5 }
        '3' { 30 }
        '4' { Read-PMCustomSeconds }
        default { 10 }
    }
    if ($interval -le 0) { return $null }

    Write-Host ''
    Write-Host '  For how long?' -ForegroundColor Cyan
    Write-PMMenuOption -Key '1' -Text '1 minute'
    Write-PMMenuOption -Key '2' -Text '5 minutes' -Default
    Write-PMMenuOption -Key '3' -Text '10 minutes'
    Write-PMMenuOption -Key '4' -Text 'Custom duration...'
    Write-Host ''
    $dc = Read-Host '  Choice [2]'
    if ([string]::IsNullOrWhiteSpace($dc)) { $dc = '2' }
    $duration = switch ($dc.Trim()) {
        '1' { 1 }
        '3' { 10 }
        '4' { Read-PMCustomMinutes }
        default { 5 }
    }
    if ($duration -le 0) { return $null }

    return @{ IntervalSeconds = $interval; DurationMinutes = $duration }
}

function Test-PMAdHocSmbOnce {
    param([Parameter(Mandatory)][object]$Target, [int]$TimeoutMs = 1500)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $portOk = Test-PMSmbPort -ComputerName $Target.Server -Port $Target.Port -TimeoutMs $TimeoutMs
    $sw.Stop()
    $latency = if ($portOk) { [int]$sw.ElapsedMilliseconds } else { $null }

    # Auth (if any) already happened once, up front, in Invoke-PMAdHocSmbRun -
    # this just re-checks the path each tick over that same session, exactly
    # like testing as the current account needs no mapping at all.
    $shareOk = $null
    if ($portOk -and -not [string]::IsNullOrWhiteSpace($Target.SharePath)) {
        try { $shareOk = [bool](Test-Path -LiteralPath $Target.SharePath -ErrorAction Stop) }
        catch { $shareOk = $false }
    }

    return [pscustomobject]@{ PortOk = $portOk; ShareOk = $shareOk; LatencyMs = $latency }
}

function Write-PMAdHocResult {
    param([Parameter(Mandatory)][object]$Target, [Parameter(Mandatory)][object]$Result)

    Write-Host ''
    if (-not $Result.PortOk) {
        Write-Host ('  Port {0}: UNREACHABLE' -f $Target.Port) -ForegroundColor Red
        return
    }
    Write-Host ('  Port {0}: reachable ({1} ms)' -f $Target.Port, $Result.LatencyMs) -ForegroundColor Green
    if ($null -ne $Result.ShareOk) {
        if ($Result.ShareOk) { Write-Host ('  Share {0}: accessible' -f $Target.SharePath) -ForegroundColor Green }
        else                 { Write-Host ('  Share {0}: NOT accessible' -f $Target.SharePath) -ForegroundColor Red }
    }
}

function Write-PMAdHocTick {
    param([Parameter(Mandatory)][string]$Stamp, [Parameter(Mandatory)][object]$Target, [Parameter(Mandatory)][object]$Result)

    if (-not $Result.PortOk) {
        Write-Host ("  [{0}] FAIL - port {1} unreachable" -f $Stamp, $Target.Port) -ForegroundColor Red
    }
    elseif ($null -ne $Result.ShareOk -and -not $Result.ShareOk) {
        Write-Host ("  [{0}] FAIL - share not accessible ({1} ms port)" -f $Stamp, $Result.LatencyMs) -ForegroundColor Red
    }
    else {
        Write-Host ("  [{0}] OK   - {1} ms" -f $Stamp, $Result.LatencyMs) -ForegroundColor Green
    }
}

# Anchors each tick to the clock rather than sleeping a fixed amount, same
# reasoning as Start-PMMonitor.ps1's sampling loop - so the time a slow tick
# takes does not make the series drift past the requested duration. Ctrl+C
# still prints the summary for whatever ran before it, via the same
# try/finally shape Start-PMMonitor.ps1 uses to keep partial results.
function Start-PMAdHocSmbLoop {
    param([Parameter(Mandatory)][object]$Target, [Parameter(Mandatory)][int]$IntervalSeconds, [Parameter(Mandatory)][int]$DurationMinutes)

    $endAt     = (Get-Date).AddMinutes($DurationMinutes)
    $attempt   = 0
    $success   = 0
    $latencies = New-Object System.Collections.Generic.List[double]

    Write-Host ''
    Write-Host ("  Testing every {0}s until about {1}. Ctrl+C to stop early and still see the summary." -f `
        $IntervalSeconds, $endAt.ToString('HH:mm:ss')) -ForegroundColor DarkGray
    Write-Host ''

    try {
        while ((Get-Date) -lt $endAt) {
            $tickStart = Get-Date
            $result    = Test-PMAdHocSmbOnce -Target $Target
            $attempt++

            $ok = $result.PortOk -and ($null -eq $result.ShareOk -or $result.ShareOk)
            if ($ok) { $latencies.Add([double]$result.LatencyMs) | Out-Null; $success++ }
            Write-PMAdHocTick -Stamp $tickStart.ToString('HH:mm:ss') -Target $Target -Result $result

            $remaining = $IntervalSeconds - ((Get-Date) - $tickStart).TotalSeconds
            if ($remaining -gt 0) {
                if ((Get-Date).AddSeconds($remaining) -ge $endAt) { break }
                Start-Sleep -Milliseconds ([int]($remaining * 1000))
            }
        }
    }
    finally {
        Write-Host ''
        if ($attempt -eq 0) {
            Write-Host '  No attempts completed.' -ForegroundColor Yellow
        }
        else {
            $rate  = [math]::Round(($success / $attempt) * 100, 1)
            $color = if ($success -eq $attempt) { 'Green' } else { 'Yellow' }
            Write-Host ("  {0}/{1} succeeded ({2}%)" -f $success, $attempt, $rate) -ForegroundColor $color
            if ($latencies.Count -gt 0) {
                $avg = [math]::Round(($latencies | Measure-Object -Average).Average, 0)
                $min = [math]::Round(($latencies | Measure-Object -Minimum).Minimum, 0)
                $max = [math]::Round(($latencies | Measure-Object -Maximum).Maximum, 0)
                Write-Host ("  Latency avg {0} ms, min {1} ms, max {2} ms" -f $avg, $min, $max) -ForegroundColor DarkGray
            }
        }
    }
}

# Authenticates once (if credentials were given) rather than per tick: on a
# 10 s interval a 10-minute repeat test is 60 attempts, and re-authenticating
# every single one is both slow and, with the wrong password, a fast way to
# trip an account lockout policy. One mapping up front, reused by every
# Test-Path call in the loop, removed once at the end - same "leaves nothing
# behind" rule the SMBCONN check itself follows.
function Invoke-PMAdHocSmbRun {
    param(
        [Parameter(Mandatory)][object]$Target,
        [int]$IntervalSeconds = 0,
        [int]$DurationMinutes = 0
    )

    $needsAuth = ($null -ne $Target.Credential) -and (-not [string]::IsNullOrWhiteSpace($Target.SharePath))
    $mapped    = $false

    if ($needsAuth) {
        Write-Host ''
        Write-Host ('  Authenticating to {0} as {1}...' -f $Target.Server, $Target.Credential.UserName) -ForegroundColor DarkGray
        try {
            New-SmbMapping -RemotePath $Target.SharePath -Credential $Target.Credential -Persistent $false -ErrorAction Stop | Out-Null
            $mapped = $true
            Write-Host '  Authenticated.' -ForegroundColor Green
        }
        catch {
            Write-Host ('  Authentication FAILED: {0}' -f $_.Exception.Message) -ForegroundColor Red
            Write-Host ''
            Read-Host '  Press Enter to continue' | Out-Null
            return
        }
    }

    try {
        if ($IntervalSeconds -gt 0 -and $DurationMinutes -gt 0) {
            Start-PMAdHocSmbLoop -Target $Target -IntervalSeconds $IntervalSeconds -DurationMinutes $DurationMinutes
        }
        else {
            $result = Test-PMAdHocSmbOnce -Target $Target
            Write-PMAdHocResult -Target $Target -Result $result
        }
    }
    finally {
        if ($mapped) { Remove-SmbMapping -RemotePath $Target.SharePath -Force -Confirm:$false -ErrorAction SilentlyContinue }
    }

    Write-Host ''
    Read-Host '  Press Enter to continue' | Out-Null
}

function Show-PMNasTestMenu {
    while ($true) {
        Clear-Host
        Write-PMRule
        Write-Host '  Test NAS / SMB connectivity' -ForegroundColor Cyan
        Write-PMRule
        Write-Host ''

        # Read fresh each pass, same reasoning as $agsHint on the main screen.
        $smbTargets = @(Get-PMSetting -Path 'SmbConnectivity.Targets' -Default @())
        if ($smbTargets.Count -gt 0) {
            Write-Host ('  Configured targets: {0}' -f $smbTargets.Count) -ForegroundColor Green
        }
        else {
            Write-Host '  Configured targets: none (Config\settings.json, SmbConnectivity.Targets)' -ForegroundColor Yellow
        }
        Write-Host ''

        if ($smbTargets.Count -gt 0) {
            Write-PMMenuOption -Key '1' -Text 'Test configured targets now (adds to the HTML report)' -Default
        }
        Write-PMMenuOption -Key '2' -Text 'Test a custom path now (console only, nothing saved)'
        Write-Host ''
        Write-PMMenuOption -Key 'B' -Text 'Back to the main menu'
        Write-Host ''

        $default = if ($smbTargets.Count -gt 0) { '1' } else { '2' }
        $c = Read-Host ('  Choice [{0}]' -f $default)
        if ([string]::IsNullOrWhiteSpace($c)) { $c = $default }

        switch ($c.Trim().ToUpper()) {
            '1' {
                if ($smbTargets.Count -gt 0) {
                    & (Join-Path $PMRoot 'Start-PMCheck.ps1') -Only SMBCONN -OutputRoot $OutputRoot -ConfigDir $ConfigDir -OpenReport
                    exit $LASTEXITCODE
                }
                else {
                    Write-Host ''
                    Write-Host '  Not a valid choice.' -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                }
            }
            '2' {
                $target = Read-PMAdHocSmbTarget
                if ($null -ne $target) {
                    $schedule = Read-PMTestSchedule
                    if ($null -eq $schedule) { Invoke-PMAdHocSmbRun -Target $target }
                    else { Invoke-PMAdHocSmbRun -Target $target -IntervalSeconds $schedule.IntervalSeconds -DurationMinutes $schedule.DurationMinutes }
                }
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

        # SmbConnectivity.Targets has no wizard of its own - it is a plain
        # array in Config\settings.json, not a credential that needs DPAPI
        # storage like the ArcGIS connection - so the hint just reflects
        # whatever is already there.
        $smbTargets = @(Get-PMSetting -Path 'SmbConnectivity.Targets' -Default @())
        if ($smbTargets.Count -gt 0) { $smbHint = ('({0} target(s) configured)' -f $smbTargets.Count) }
        else                         { $smbHint = '(not configured)' }

        Write-Host '  What do you want to check?'
        Write-Host ''
        Write-PMMenuOption -Key '1' -Text 'Server - the regular maintenance checks' -Default
        Write-PMMenuOption -Key '2' -Text 'ArcGIS Server - site, services, usage reports   ' -NoNewline
        Write-Host $agsHint -ForegroundColor DarkGray
        Write-PMMenuOption -Key '3' -Text 'ArcGIS Server - service & database connection   ' -NoNewline
        Write-Host $agsHint -ForegroundColor DarkGray
        Write-Host ''
        Write-PMMenuOption -Key '4' -Text 'Test NAS / SMB connectivity                     ' -NoNewline
        Write-Host $smbHint -ForegroundColor DarkGray
        Write-Host ''
        Write-PMMenuOption -Key 'A' -Text 'ArcGIS Server connection...'
        Write-Host ''
        Write-PMMenuOption -Key 'Q' -Text 'Quit'
        Write-Host ''

        $modeChoice = Read-Host '  Choice [1]'
        if ([string]::IsNullOrWhiteSpace($modeChoice)) { $modeChoice = '1' }

        switch ($modeChoice.Trim().ToUpper()) {
            '1' { $mode = 'Server' }
            '2' { if (Confirm-PMArcGISConnectionReady) { $mode = 'ArcGIS' } }
            '3' { if (Confirm-PMArcGISConnectionReady) { $mode = 'ArcGISDataConn' } }
            '4' { Show-PMNasTestMenu }
            'A' { Show-PMArcGISMenu }
            'Q' { Write-Host ''; Write-Host '  Cancelled.'; Write-Host ''; exit 0 }
            default {
                Write-Host ''
                Write-Host '  Not a valid choice.' -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
        }
    }

    if ($mode -eq 'ArcGIS' -or $mode -eq 'ArcGISDataConn') { $ready = $true; continue }

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
        Write-PMMenuOption -Key '1' -Text 'No sampling - assess now                 ' -NoNewline
        Write-Host '(about 10 s, default)' -ForegroundColor DarkGray
        Write-PMMenuOption -Key '2' -Text ('Sample 15 minutes, then assess           ({0})' -f (Get-PMFinishText 15))
        Write-PMMenuOption -Key '3' -Text ('Sample {0} minutes, then assess           ({1})' -f $defaultMinutes, (Get-PMFinishText $defaultMinutes))
        Write-PMMenuOption -Key '4' -Text ('Sample 60 minutes, then assess           ({0})' -f (Get-PMFinishText 60))
        Write-PMMenuOption -Key '5' -Text 'Sample for a custom duration...'
        Write-PMMenuOption -Key '6' -Text 'Sample only - do not build a report'
        Write-Host ''
        Write-PMMenuOption -Key 'B' -Text 'Back'
        Write-PMMenuOption -Key 'Q' -Text 'Quit'
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

if ($mode -eq 'ArcGISDataConn') {
    # Service & database connection only: AGSSVC (map service status),
    # AGSDATA (can each registered data store still connect) and
    # AGSWORKSPACE (what workspace - server/database or path - each
    # service actually connects to, Referenced/Replaced/Copied) - not
    # AGS/AGSUSAGE/AGSLOG, so -Only is used instead of -Group ArcGIS. An
    # explicit -Only already bypasses Checks.Disabled the same way -Group
    # does (see Start-PMCheck.ps1).
    & (Join-Path $PMRoot 'Start-PMCheck.ps1') -Only 'AGSSVC,AGSDATA,AGSWORKSPACE' -OutputRoot $OutputRoot -ConfigDir $ConfigDir -OpenReport
    exit $LASTEXITCODE
}

# ArcGIS mode: no sampling question, nothing to trend - just run every
# ArcGIS check against the connection confirmed to exist above.
& (Join-Path $PMRoot 'Start-PMCheck.ps1') -Group ArcGIS -OutputRoot $OutputRoot -ConfigDir $ConfigDir -OpenReport
exit $LASTEXITCODE
