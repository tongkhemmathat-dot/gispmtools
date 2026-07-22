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
Initialize-PMCore -ConfigDir $ConfigDir

$defaultMinutes = [int](Get-PMSetting -Path 'Monitor.DefaultMinutes'  -Default 30)
$interval       = [int](Get-PMSetting -Path 'Monitor.IntervalSeconds' -Default 10)
$maxAgeHours    = [double](Get-PMSetting -Path 'Monitor.MaxDataAgeHours' -Default 24)
$toolVersion    = [string](Get-PMSetting -Path 'Report.ToolVersion' -Default '1.1.0')

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

# --- the menu ---------------------------------------------------------------
$minutes   = 0
$runReport = $true

while ($true) {

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
    Write-Host '   [Q]  Quit'
    Write-Host ''

    $choice = Read-Host '  Choice [1]'
    if ([string]::IsNullOrWhiteSpace($choice)) { $choice = '1' }

    switch ($choice.Trim().ToUpper()) {
        '1' { $minutes = 0;               $runReport = $true;  break }
        '2' { $minutes = 15;              $runReport = $true;  break }
        '3' { $minutes = $defaultMinutes; $runReport = $true;  break }
        '4' { $minutes = 60;              $runReport = $true;  break }
        '5' {
            $custom = Read-PMCustomMinutes
            if ($custom -eq 0) { continue }
            $minutes = $custom; $runReport = $true; break
        }
        '6' {
            $custom = Read-PMCustomMinutes
            if ($custom -eq 0) { continue }
            $minutes = $custom; $runReport = $false; break
        }
        'Q' { Write-Host ''; Write-Host '  Cancelled.'; Write-Host ''; exit 0 }
        default {
            Write-Host ''
            Write-Host '  Not a valid choice.' -ForegroundColor Yellow
            Start-Sleep -Seconds 1
            continue
        }
    }
    break
}

# --- act on it --------------------------------------------------------------
Write-Host ''

if ($minutes -gt 0) {
    & (Join-Path $PMRoot 'Start-PMMonitor.ps1') `
        -Minutes $minutes -IntervalSeconds $interval -OutputRoot $OutputRoot -ConfigDir $ConfigDir

    if (-not $runReport) {
        Write-Host 'Sampling finished. Run Run-PM.cmd again and choose 1 to build the report.'
        Write-Host ''
        exit 0
    }
}

& (Join-Path $PMRoot 'Start-PMCheck.ps1') -OutputRoot $OutputRoot -ConfigDir $ConfigDir -OpenReport
exit $LASTEXITCODE
