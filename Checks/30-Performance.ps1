# PMtools check - CPU and memory utilisation.  ASCII-only; text comes from i18n.json.
#
# A spot reading taken at the moment of the assessment. For sustained load see
# the TREND check, which charts samples collected by Start-PMMonitor.ps1; both
# read through the same Core helpers so the two sections cannot disagree.

function Invoke-PMCheckPerformance {

    # A single reading is easily unrepresentative, so take two a second apart
    # and average them.
    $samples = @()
    for ($i = 0; $i -lt 2; $i++) {
        $load = Get-PMCpuPercent
        if ($null -ne $load) { $samples += [double]$load }
        if ($i -eq 0) { Start-Sleep -Seconds 1 }
    }
    if ($samples.Count -gt 0) { $cpuPct = [math]::Round(($samples | Measure-Object -Average).Average, 1) }
    else                      { $cpuPct = 0 }

    $mem     = Get-PMMemoryUsage
    $totalGB = $mem.TotalGB
    $freeGB  = $mem.FreeGB
    $usedGB  = $mem.UsedGB
    $freePct = $mem.FreePercent

    $cpuStatus = Test-PMThreshold -Name 'CpuPercent'     -Value $cpuPct
    $memStatus = Test-PMThreshold -Name 'MemFreePercent' -Value $freePct

    $rows = @(
        (New-PMItemRow -TextKey 'perf.item.cpu'        -Value ("{0} %"  -f $cpuPct)  -Status $cpuStatus)
        (New-PMItemRow -TextKey 'perf.item.ramtotal'   -Value ("{0} GB" -f $totalGB))
        (New-PMItemRow -TextKey 'perf.item.ramused'    -Value ("{0} GB" -f $usedGB))
        (New-PMItemRow -TextKey 'perf.item.ramfree'    -Value ("{0} GB" -f $freeGB))
        (New-PMItemRow -TextKey 'perf.item.ramfreepct' -Value ("{0} %"  -f $freePct) -Status $memStatus)
    )

    $page = @(Get-CimInstance -ClassName Win32_PageFileUsage -ErrorAction SilentlyContinue)
    if ($page.Count -gt 0) {
        $alloc = ($page | Measure-Object -Property AllocatedBaseSize -Sum).Sum
        $inUse = ($page | Measure-Object -Property CurrentUsage -Sum).Sum
        $rows += New-PMItemRow -TextKey 'perf.item.pagefile' -Value ("{0} MB / {1} MB" -f $inUse, $alloc)
    }

    $findings = @()
    if ($cpuStatus -eq 'WARN' -or $cpuStatus -eq 'CRIT') {
        $findings += New-PMFinding -Severity $cpuStatus -TextKey 'perf.finding.cpu' -Values @($cpuPct)
    }
    if ($memStatus -eq 'WARN' -or $memStatus -eq 'CRIT') {
        $findings += New-PMFinding -Severity $memStatus -TextKey 'perf.finding.mem' -Values @($freePct, $freeGB)
    }

    return New-PMResult -Id 'PERF' -TitleKey 'perf.title' `
        -Status (Get-PMWorstStatus @($cpuStatus, $memStatus)) `
        -SummaryKey 'perf.summary' -SummaryValues @($cpuPct, $freePct, $freeGB) `
        -Columns (New-PMItemColumns) -Rows $rows -Findings $findings `
        -Raw ([pscustomobject]@{
            CpuPercent = $cpuPct; CpuSamples = $samples
            MemoryTotalGB = $totalGB; MemoryUsedGB = $usedGB
            MemoryFreeGB = $freeGB; MemoryFreePercent = $freePct
        })
}

Register-PMCheck -Id 'PERF' -TitleKey 'perf.title' -Function 'Invoke-PMCheckPerformance'
