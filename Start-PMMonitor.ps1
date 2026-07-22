<#
    PMtools - CPU and memory sampler.

    Records one CPU and memory reading every few seconds for a fixed period and
    writes them to a CSV. The PERFTREND check picks up the newest CSV and draws
    the trend chart into the report.

    READ-ONLY, like the rest of the tool: it queries counters and writes one CSV
    under Output\_Perf. Nothing on the server is changed. Each sample costs about
    0.4 s of work, so at the default ten-second interval the sampler uses roughly
    4% of one core - it observes the load without meaningfully adding to it.

    ASCII-only; see README.md for why.

    Examples
        .\Start-PMMonitor.ps1                       30 minutes, every 10 s
        .\Start-PMMonitor.ps1 -Minutes 60           one hour
        .\Start-PMMonitor.ps1 -Minutes 5 -IntervalSeconds 2
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 1440)][int]$Minutes = 30,

    [ValidateRange(1, 300)][int]$IntervalSeconds = 10,

    # Left empty here on purpose - see the note below the param block.
    [string]$OutputRoot,

    [string]$ConfigDir,

    # Run the assessment and build the report as soon as sampling finishes.
    [switch]$ThenReport
)

$ErrorActionPreference = 'Stop'

# $PSScriptRoot is empty while param() defaults are evaluated when a
# [CmdletBinding()] script runs under `powershell.exe -File` on PS 5.1.
# Resolve the folder in the body, where it is reliable. See Start-PMCheck.ps1.
$PMRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($PMRoot)) { $PMRoot = $PSScriptRoot }

if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Join-Path $PMRoot 'Output' }
if ([string]::IsNullOrWhiteSpace($ConfigDir))  { $ConfigDir  = Join-Path $PMRoot 'Config' }

. (Join-Path $PMRoot 'Lib\Core.ps1')
Initialize-PMCore -ConfigDir $ConfigDir

$startedAt = Get-Date
$endAt     = $startedAt.AddMinutes($Minutes)
$total     = [int](($Minutes * 60) / $IntervalSeconds)

$perfDir = Join-Path $OutputRoot '_Perf'
if (-not (Test-Path -LiteralPath $perfDir)) { New-Item -ItemType Directory -Path $perfDir -Force | Out-Null }

$csvPath = Join-Path $perfDir ("{0}-perf-{1}.csv" -f $env:COMPUTERNAME, $startedAt.ToString('yyyyMMdd-HHmm'))

Write-PMLog ""
Write-PMLog "PMtools - CPU and memory sampling" -Level Step
Write-PMLog ("Server   : {0}" -f $env:COMPUTERNAME)
Write-PMLog ("Duration : {0} minutes, one sample every {1} s ({2} samples)" -f $Minutes, $IntervalSeconds, $total)
Write-PMLog ("Finishes : {0}" -f $endAt.ToString('yyyy-MM-dd HH:mm:ss'))
Write-PMLog ("File     : {0}" -f $csvPath)
Write-PMLog ""
Write-PMLog "Leave this window open. Press Ctrl+C to stop early - samples already"
Write-PMLog "taken are kept and the report will use them."
Write-PMLog ""

$samples = New-Object System.Collections.ArrayList
$taken   = 0
$failed  = 0

try {
    while ((Get-Date) -lt $endAt) {

        # Anchor the next tick to the clock rather than sleeping a fixed amount,
        # so the time taken by a sample does not make the series drift.
        $tickStart = Get-Date

        try {
            $cpu = Get-PMCpuPercent
            $mem = Get-PMMemoryUsage

            if ($null -ne $cpu -and $null -ne $mem) {
                [void]$samples.Add([pscustomobject]@{
                    Timestamp      = $tickStart.ToString('yyyy-MM-dd HH:mm:ss')
                    CpuPercent     = [math]::Round($cpu, 1)
                    MemUsedPercent = $mem.UsedPercent
                    MemUsedGB      = $mem.UsedGB
                    MemTotalGB     = $mem.TotalGB
                })
                $taken++
            }
            else { $failed++ }
        }
        catch { $failed++ }   # one bad reading must not end a 30-minute run

        # Flush after every sample: if the window is closed or the server is
        # restarted mid-run, everything collected so far is already on disk.
        $samples | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

        $elapsed = ((Get-Date) - $startedAt).TotalSeconds
        $pct     = [math]::Min(100, [math]::Round(($elapsed / ($Minutes * 60)) * 100, 0))
        $lastCpu = 0
        $lastMem = 0
        if ($samples.Count -gt 0) {
            $lastCpu = $samples[$samples.Count - 1].CpuPercent
            $lastMem = $samples[$samples.Count - 1].MemUsedPercent
        }
        Write-Progress -Activity 'PMtools sampling' `
            -Status ("{0} samples | CPU {1}% | Memory {2}% | finishes {3}" -f $taken, $lastCpu, $lastMem, $endAt.ToString('HH:mm')) `
            -PercentComplete $pct

        $remaining = $IntervalSeconds - ((Get-Date) - $tickStart).TotalSeconds
        if ($remaining -gt 0) {
            if ((Get-Date).AddSeconds($remaining) -gt $endAt) { break }
            Start-Sleep -Milliseconds ([int]($remaining * 1000))
        }
    }
}
finally {
    Write-Progress -Activity 'PMtools sampling' -Completed
    if ($samples.Count -gt 0) {
        $samples | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
    }
}

Write-PMLog ""
if ($taken -eq 0) {
    Write-PMLog "No samples were collected. Check that the performance counters are working." -Level Bad
    exit 3
}

$cpuValues = @($samples | ForEach-Object { [double]$_.CpuPercent })
$memValues = @($samples | ForEach-Object { [double]$_.MemUsedPercent })

Write-PMLog ("Collected {0} samples over {1:N1} minutes" -f $taken, ((Get-Date) - $startedAt).TotalMinutes) -Level Good
if ($failed -gt 0) { Write-PMLog ("  {0} reading(s) failed and were skipped" -f $failed) -Level Warn }
Write-PMLog ("  CPU    avg {0,5:N1}%  p95 {1,5:N1}%  max {2,5:N1}%" -f `
    ($cpuValues | Measure-Object -Average).Average, (Get-PMPercentile -Values $cpuValues), ($cpuValues | Measure-Object -Maximum).Maximum)
Write-PMLog ("  Memory avg {0,5:N1}%  p95 {1,5:N1}%  max {2,5:N1}%" -f `
    ($memValues | Measure-Object -Average).Average, (Get-PMPercentile -Values $memValues), ($memValues | Measure-Object -Maximum).Maximum)
Write-PMLog ""
Write-PMLog ("Saved: {0}" -f $csvPath) -Level Good
Write-PMLog ""

if ($ThenReport) {
    Write-PMLog "Running the assessment..." -Level Step
    & (Join-Path $PMRoot 'Start-PMCheck.ps1') -OutputRoot $OutputRoot -ConfigDir $ConfigDir -OpenReport
}
else {
    Write-PMLog "Now run .\Start-PMCheck.ps1 to build the report with this trend included."
}
