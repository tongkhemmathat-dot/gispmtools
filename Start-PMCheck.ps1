<#
    PMtools - Preventive Maintenance assessment for Windows Server.

    Copy the whole PMtools folder onto the server, then either double-click
    Run-PM.cmd or run this script from an elevated PowerShell prompt.

    READ-ONLY BY DESIGN. Every check queries state and nothing else: no setting
    is changed, no service is started or stopped, nothing is installed, and the
    server is never restarted. The only thing written is the report folder under
    Output\. The one check that reaches outside the machine - WU, which asks
    WSUS or Microsoft Update what is pending - is disabled by default in
    Config\settings.json because that search is slow and load-bearing on the
    update infrastructure; run it deliberately with -Only WU if it is wanted.

    ASCII-only by design: PowerShell 5.1 reads .ps1 files as ANSI unless they
    carry a UTF-8 BOM, so all Thai text lives in Config\i18n.json instead.

    Examples
        .\Start-PMCheck.ps1
        .\Start-PMCheck.ps1 -OpenReport
        .\Start-PMCheck.ps1 -Only DISK,CERT
        .\Start-PMCheck.ps1 -Skip WU
#>
[CmdletBinding()]
param(
    # Where the timestamped result folder is created.
    # Left empty here on purpose - see the note below the param block.
    [string]$OutputRoot,

    [string]$ConfigDir,

    # Run only these check ids (e.g. DISK, CERT). Default: all registered checks.
    [string[]]$Only,

    # Run everything except these check ids. Useful for -Skip WU, which needs
    # to reach an update source and can be slow on isolated networks.
    [string[]]$Skip,

    # Sample CPU and memory for this many minutes BEFORE running the checks, so
    # the report carries a trend chart rather than a single spot reading.
    # Equivalent to running Start-PMMonitor.ps1 first.
    [ValidateRange(0, 1440)][int]$MonitorMinutes = 0,

    # Open the finished report in the default browser.
    [switch]$OpenReport
)

$ErrorActionPreference = 'Stop'
$startedAt = Get-Date

# On PowerShell 5.1, a script carrying [CmdletBinding()] has an EMPTY
# $PSScriptRoot while its param() default expressions are evaluated, but only
# when it is launched as `powershell.exe -File script.ps1` - which is exactly
# how Run-PM.cmd starts it. Defaults built from $PSScriptRoot up there therefore
# blew up on the one path real users take, while `& .\script.ps1` from an open
# PowerShell session worked fine and hid the fault. Resolve the folder here in
# the body instead, where it is reliable, and default the paths afterwards.
$PMRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($PMRoot)) { $PMRoot = $PSScriptRoot }

if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Join-Path $PMRoot 'Output' }
if ([string]::IsNullOrWhiteSpace($ConfigDir))  { $ConfigDir  = Join-Path $PMRoot 'Config' }

# --- load the shared contract and the renderer ------------------------------
. (Join-Path $PMRoot 'Lib\Core.ps1')
. (Join-Path $PMRoot 'Lib\Report.Html.ps1')
# Needed by the Checks\A*-ArcGIS*.ps1 checks. Loaded unconditionally because
# a check file is dot-sourced before anything knows whether it will run, and
# this file only defines functions - it opens no connection by itself.
. (Join-Path $PMRoot 'Lib\ArcGIS.ps1')

Initialize-PMCore -ConfigDir $ConfigDir
Set-PMOutputRoot -Path $OutputRoot

$toolVersion = [string](Get-PMSetting -Path 'Report.ToolVersion' -Default '1.2.0')

Write-PMLog ""
Write-PMLog "PMtools $toolVersion - Preventive Maintenance assessment" -Level Step
Write-PMLog ("Server : {0}" -f $env:COMPUTERNAME)
Write-PMLog ("Started: {0}" -f $startedAt.ToString('yyyy-MM-dd HH:mm:ss'))
Write-PMLog ""

# Several checks read machine-wide state that is only visible to an admin.
# Warn loudly but keep going, and record the fact in the report itself.
$isAdmin = Test-PMIsAdministrator
if (-not $isAdmin) {
    Write-PMLog ("WARNING: {0}" -f (Get-PMText -Key 'ui.notAdminWarning').En) -Level Warn
    Write-PMLog ""
}

# --- optional sampling pass -------------------------------------------------
# Runs first so the TREND check finds fresh data when the checks execute.
if ($MonitorMinutes -gt 0) {
    $interval = [int](Get-PMSetting -Path 'Monitor.IntervalSeconds' -Default 10)
    & (Join-Path $PMRoot 'Start-PMMonitor.ps1') `
        -Minutes $MonitorMinutes -IntervalSeconds $interval -OutputRoot $OutputRoot -ConfigDir $ConfigDir
    Write-PMLog ""
}

# --- discover checks --------------------------------------------------------
Clear-PMRegisteredCheck
$checkFiles = @(Get-ChildItem -LiteralPath (Join-Path $PMRoot 'Checks') -Filter '*.ps1' -ErrorAction Stop |
                Sort-Object Name)
foreach ($file in $checkFiles) {
    try { . $file.FullName }
    catch { Write-PMLog ("Failed to load check file {0}: {1}" -f $file.Name, $_.Exception.Message) -Level Bad }
}

$checks = @(Get-PMRegisteredCheck)

# powershell.exe -File hands every argument over as a literal string, so
# "-Only SYSTEM,DISK" arrives as ONE element "SYSTEM,DISK" rather than the
# three-element array the same switch produces when the script is called as
# .\Start-PMCheck.ps1 from an open prompt. Nothing then matches an id and the
# run dies with "No checks were selected", which points at the wrong thing
# entirely. Splitting here makes both launch paths behave the same; check ids
# never contain a comma, so this cannot merge two real ids by accident.
#
# The single-file bootstrap in Build-PMSingle.ps1 already did this for its own
# forwarded arguments - this is the same trap one layer down, and it is the
# third time in this project that -File and & have differed. See HANDOVER.md.
function Expand-PMIdList {
    param([string[]]$Values)
    if (-not $Values) { return @() }
    return @($Values |
        ForEach-Object { $_ -split ',' } |
        ForEach-Object { $_.Trim() } |
        Where-Object   { $_ })
}

$Only = Expand-PMIdList -Values $Only
$Skip = Expand-PMIdList -Values $Skip

if ($Only) {
    # An explicit -Only is the operator asking for exactly these, so it also
    # overrides the disabled list in settings.json.
    $checks = @($checks | Where-Object { $Only -contains $_.Id })
}
else {
    $disabled = @(Get-PMSetting -Path 'Checks.Disabled' -Default @())
    if ($disabled.Count -gt 0) {
        $checks = @($checks | Where-Object { $disabled -notcontains $_.Id })
        Write-PMLog ("Disabled in settings.json, not run: {0}" -f ($disabled -join ', '))
        Write-PMLog ""
    }
}

if ($Skip) { $checks = @($checks | Where-Object { $Skip -notcontains $_.Id }) }

if ($checks.Count -eq 0) {
    throw "No checks were selected. Available ids: $((Get-PMRegisteredCheck | ForEach-Object Id) -join ', ')"
}

# --- run them ---------------------------------------------------------------
$results = @()
$index   = 0

foreach ($check in $checks) {
    $index++
    $title = (Get-PMText -Key $check.TitleKey).En
    Write-Progress -Activity 'PMtools' -Status $title -PercentComplete (($index / $checks.Count) * 100)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        # A check that throws must not take the whole report down: it degrades
        # to a visible ERROR card carrying the real exception message.
        $result = & $check.Function
        if ($null -eq $result) { throw "The check returned no result." }
    }
    catch {
        $result = New-PMErrorResult -Id $check.Id -TitleKey $check.TitleKey -Message $_.Exception.Message
    }
    $sw.Stop()
    $result.DurationMs = [int]$sw.ElapsedMilliseconds

    switch ($result.Status) {
        'CRIT'  { $level = 'Bad' }
        'ERROR' { $level = 'Bad' }
        'WARN'  { $level = 'Warn' }
        'OK'    { $level = 'Good' }
        default { $level = 'Info' }
    }
    Write-PMLog ("  [{0,-5}] {1,-42} {2,6} ms" -f $result.Status, $title, $result.DurationMs) -Level $level

    $results += $result
}
Write-Progress -Activity 'PMtools' -Completed

# --- write the output folder ------------------------------------------------
$finishedAt  = Get-Date
$durationSec = [math]::Round(($finishedAt - $startedAt).TotalSeconds, 1)

$outDir = Join-Path $OutputRoot ("{0}_{1}" -f $env:COMPUTERNAME, $startedAt.ToString('yyyyMMdd-HHmm'))
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$meta = [pscustomobject]@{
    Hostname    = $env:COMPUTERNAME
    GeneratedAt = $finishedAt
    DurationSec = $durationSec
    ToolVersion = $toolVersion
    IsAdmin     = $isAdmin
    Config      = Get-PMConfig
}

$jsonPath = Join-Path $outDir 'PM-Data.json'
$payload  = [pscustomobject]@{
    Meta    = [pscustomobject]@{
        Hostname    = $meta.Hostname
        GeneratedAt = $meta.GeneratedAt
        DurationSec = $meta.DurationSec
        ToolVersion = $meta.ToolVersion
        IsAdmin     = $meta.IsAdmin
    }
    Results = $results
}
# ConvertTo-Json escapes non-ASCII to \uXXXX, which is valid JSON and keeps the
# Thai text readable to any consumer regardless of how the file is opened.
$payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

$htmlPath = Join-Path $outDir 'PM-Report.html'
$html     = New-PMHtmlReport -Results $results -Meta $meta
# UTF-8 *with* BOM so browsers never guess the encoding wrong on the Thai text.
[System.IO.File]::WriteAllText($htmlPath, $html, (New-Object System.Text.UTF8Encoding($true)))

# --- console summary --------------------------------------------------------
$counts = @{ OK = 0; WARN = 0; CRIT = 0; INFO = 0; ERROR = 0 }
foreach ($r in $results) { $counts[$r.Status] = $counts[$r.Status] + 1 }

Write-PMLog ""
Write-PMLog ("Completed in {0} s" -f $durationSec) -Level Step
Write-PMLog ("  Normal {0} | Warning {1} | Critical {2} | Information {3} | Failed {4}" -f `
    $counts.OK, $counts.WARN, $counts.CRIT, $counts.INFO, $counts.ERROR)
Write-PMLog ""
Write-PMLog ("Report : {0}" -f $htmlPath) -Level Good
Write-PMLog ("Data   : {0}" -f $jsonPath)
Write-PMLog ""

if ($OpenReport) { Start-Process $htmlPath }

# Non-zero exit when something needs attention, so the tool can be wired into a
# scheduled task or monitoring wrapper later without changing anything here.
if ($counts.CRIT -gt 0 -or $counts.ERROR -gt 0) { exit 2 }
elseif ($counts.WARN -gt 0) { exit 1 }
else { exit 0 }
