# PMtools check - CPU and memory trend over time.  ASCII-only; text from i18n.json.
#
# Reads the newest sample file written by Start-PMMonitor.ps1. The PERF check
# above it is a spot reading taken at the moment of the assessment; this one
# answers the question a spot reading cannot - whether the load is sustained.
#
# Status is judged on the 95th percentile rather than the maximum: a one-sample
# spike to 100% is normal on any server, a p95 of 90% is a capacity problem.

function Invoke-PMCheckPerfTrend {

    $perfDir   = Join-Path (Get-PMOutputRoot) '_Perf'
    $maxAgeHrs = [double](Get-PMSetting -Path 'Monitor.MaxDataAgeHours' -Default 24)

    $columns = @(
        (New-PMColumn -Key 'Metric' -TextKey 'trend.col.metric'),
        (New-PMColumn -Key 'Min'    -TextKey 'trend.col.min' -Align 'right'),
        (New-PMColumn -Key 'Avg'    -TextKey 'trend.col.avg' -Align 'right'),
        (New-PMColumn -Key 'P95'    -TextKey 'trend.col.p95' -Align 'right'),
        (New-PMColumn -Key 'Max'    -TextKey 'trend.col.max' -Align 'right')
    )

    # --- locate the newest sample file ---------------------------------------
    $csv = $null
    if (Test-Path -LiteralPath $perfDir) {
        $csv = Get-ChildItem -LiteralPath $perfDir -Filter '*-perf-*.csv' -ErrorAction SilentlyContinue |
               Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }

    if ($null -eq $csv) {
        return New-PMResult -Id 'TREND' -TitleKey 'trend.title' -Status 'INFO' `
            -SummaryKey 'trend.summary.none' `
            -Columns $columns -Rows @() -Findings @() `
            -Raw ([pscustomobject]@{ DataFound = $false })
    }

    $ageHours = ((Get-Date) - $csv.LastWriteTime).TotalHours
    if ($ageHours -gt $maxAgeHrs) {
        return New-PMResult -Id 'TREND' -TitleKey 'trend.title' -Status 'INFO' `
            -SummaryKey 'trend.summary.stale' -SummaryValues @([math]::Round($ageHours, 1), $maxAgeHrs) `
            -Columns $columns -Rows @() -Findings @() `
            -Raw ([pscustomobject]@{ DataFound = $false; File = $csv.Name; AgeHours = $ageHours })
    }

    $data = @(Import-Csv -LiteralPath $csv.FullName -Encoding UTF8)
    if ($data.Count -lt 2) {
        return New-PMResult -Id 'TREND' -TitleKey 'trend.title' -Status 'INFO' `
            -SummaryKey 'trend.summary.none' `
            -Columns $columns -Rows @() -Findings @() `
            -Raw ([pscustomobject]@{ DataFound = $false; SampleCount = $data.Count })
    }

    $times = @($data | ForEach-Object { [datetime]$_.Timestamp })
    $cpu   = @($data | ForEach-Object { [double]$_.CpuPercent })
    $mem   = @($data | ForEach-Object { [double]$_.MemUsedPercent })

    $spanMin = [math]::Round((($times[$times.Count - 1] - $times[0]).TotalMinutes), 0)

    # --- summary statistics ---------------------------------------------------
    $cpuStat = [pscustomobject]@{
        Min = [math]::Round((($cpu | Measure-Object -Minimum).Minimum), 1)
        Avg = [math]::Round((($cpu | Measure-Object -Average).Average), 1)
        P95 = [math]::Round((Get-PMPercentile -Values $cpu), 1)
        Max = [math]::Round((($cpu | Measure-Object -Maximum).Maximum), 1)
    }
    $memStat = [pscustomobject]@{
        Min = [math]::Round((($mem | Measure-Object -Minimum).Minimum), 1)
        Avg = [math]::Round((($mem | Measure-Object -Average).Average), 1)
        P95 = [math]::Round((Get-PMPercentile -Values $mem), 1)
        Max = [math]::Round((($mem | Measure-Object -Maximum).Maximum), 1)
    }

    $cpuStatus = Test-PMThreshold -Name 'CpuPercent' -Value $cpuStat.P95
    # MemFreePercent is expressed as free space, so convert the used figure back.
    $memStatus = Test-PMThreshold -Name 'MemFreePercent' -Value (100 - $memStat.P95)

    $cpuName = Get-PMText -Key 'trend.series.cpu'
    $memName = Get-PMText -Key 'trend.series.mem'

    $rows = @()
    foreach ($m in @(
        @{ Name = $cpuName; Stat = $cpuStat; Status = $cpuStatus },
        @{ Name = $memName; Stat = $memStat; Status = $memStatus })) {

        $row = New-PMRow -Status $m.Status -Values @{
            Metric = $m.Name.Th
            Min    = ("{0} %" -f $m.Stat.Min)
            Avg    = ("{0} %" -f $m.Stat.Avg)
            P95    = ("{0} %" -f $m.Stat.P95)
            Max    = ("{0} %" -f $m.Stat.Max)
        }
        $row['MetricEn'] = $m.Name.En
        $rows += $row
    }

    # --- chart data -----------------------------------------------------------
    # Long runs are averaged into buckets so the SVG stays small and the line
    # stays readable; a 30-minute run at the default interval is untouched.
    $maxPoints = [int](Get-PMSetting -Path 'Monitor.MaxChartPoints' -Default 360)
    $plotTimes = $times
    $plotCpu   = $cpu
    $plotMem   = $mem

    if ($data.Count -gt $maxPoints) {
        $bucket    = [math]::Ceiling($data.Count / $maxPoints)
        $plotTimes = @(); $plotCpu = @(); $plotMem = @()
        for ($i = 0; $i -lt $data.Count; $i += $bucket) {
            $end = [math]::Min($i + $bucket - 1, $data.Count - 1)
            $plotTimes += $times[$i]
            $plotCpu   += [math]::Round((($cpu[$i..$end] | Measure-Object -Average).Average), 1)
            $plotMem   += [math]::Round((($mem[$i..$end] | Measure-Object -Average).Average), 1)
        }
    }

    # Minutes alone repeat themselves on a short run ("17:00, 17:00, 17:00"),
    # so drop to seconds when the whole window is under ten minutes.
    if ($spanMin -lt 10) { $timeFormat = 'HH:mm:ss' } else { $timeFormat = 'HH:mm' }

    # Built once and shared by both languages: two separate argument lists is
    # how the English caption previously ended up one value short.
    $capValues = @($data.Count, $spanMin,
                   $times[0].ToString($timeFormat),
                   $times[$times.Count - 1].ToString($timeFormat),
                   $csv.Name)
    $caption = Get-PMText -Key 'trend.caption' -Values $capValues

    $chart = New-PMLineChart `
        -XLabels @($plotTimes | ForEach-Object { $_.ToString($timeFormat) }) `
        -Series @(
            [pscustomobject]@{ TitleTh = $cpuName.Th; TitleEn = $cpuName.En; Values = @($plotCpu) },
            [pscustomobject]@{ TitleTh = $memName.Th; TitleEn = $memName.En; Values = @($plotMem) }
        ) `
        -CaptionTh $caption.Th -CaptionEn $caption.En

    # --- findings -------------------------------------------------------------
    $findings = @()
    if ($cpuStatus -eq 'WARN' -or $cpuStatus -eq 'CRIT') {
        $findings += New-PMFinding -Severity $cpuStatus -TextKey 'trend.finding.cpu' `
            -Values @($cpuStat.P95, $spanMin, $cpuStat.Max, (Get-PMThresholdValue -Name 'CpuPercent' -Level 'Warn'))
    }
    if ($memStatus -eq 'WARN' -or $memStatus -eq 'CRIT') {
        $findings += New-PMFinding -Severity $memStatus -TextKey 'trend.finding.mem' `
            -Values @($memStat.P95, $spanMin, $memStat.Max)
    }

    $overall = Get-PMWorstStatus @($cpuStatus, $memStatus)

    return New-PMResult -Id 'TREND' -TitleKey 'trend.title' -Status $overall `
        -SummaryKey 'trend.summary' -SummaryValues @($spanMin, $data.Count, $cpuStat.Avg, $memStat.Avg) `
        -Columns $columns -Rows $rows -Findings $findings -Chart $chart `
        -Raw ([pscustomobject]@{
            DataFound   = $true
            SourceFile  = $csv.Name
            SampleCount = $data.Count
            SpanMinutes = $spanMin
            StartTime   = $times[0]
            EndTime     = $times[$times.Count - 1]
            Cpu         = $cpuStat
            Memory      = $memStat
        })
}

Register-PMCheck -Id 'TREND' -TitleKey 'trend.title' -Function 'Invoke-PMCheckPerfTrend'
