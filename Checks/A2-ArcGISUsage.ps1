# PMtools check - ArcGIS Server usage report totals (requests, response time,
# timeouts) over whatever window the site's own usage report already covers.
# ASCII-only; text comes from i18n.json.
#
# Needs a configured connection and is disabled by default, like every other
# A*-ArcGIS* check - see A0-ArcGISSite.ps1.
#
# Read-only, and stricter than the other ArcGIS checks about it: this check
# is ONLY allowed to read a usage report that already exists on the site.
# /usagereports/<name>/data can only answer for a report that has already
# been created - and creating one is a write. ArcGIS Server Manager creates
# one automatically the first time someone opens its statistics page, so
# most managed sites already have one; a site where nobody has ever opened
# that page has none, and that is reported as an INFO card, not an error.
#
# AGSSVC deliberately avoids this endpoint for the same reason and reads
# per-service instance counters instead - see its header comment. This check
# exists alongside it for the numbers only a usage report carries: total
# request volume and response time, aggregated for the whole report rather
# than per service.

function Invoke-PMCheckArcGISUsage {

    try {
        $session = Get-PMArcGISSession

        $list = Invoke-PMArcGISAdmin -Root $session.Root -Path 'usagereports' -Token $session.Token -TimeoutSec $session.TimeoutSec

        $names = @()
        foreach ($e in @($list.usagereports)) {
            if ($e -is [string])                             { $names += $e }
            elseif ($e.PSObject.Properties['reportname'])     { $names += [string]$e.reportname }
        }

        if ($names.Count -eq 0) {
            return New-PMResult -Id 'AGSUSAGE' -TitleKey 'agsusage.title' -Status 'INFO' `
                -SummaryKey 'agsusage.summary.none'
        }

        # Sites can carry more than one saved report (Manager's own "System"
        # report plus any an administrator created by hand). There is no
        # reliable way to tell which one an operator cares about from the
        # Admin API alone, so the first one listed is read - normally the
        # only one present. The report name actually used is always shown
        # in the output so this choice is never silent.
        $reportName = $names[0]
        $findings   = @()

        try {
            $resp = Invoke-PMArcGISAdmin -Root $session.Root -Path "usagereports/$reportName/data" `
                                          -Token $session.Token -TimeoutSec $session.TimeoutSec
        }
        catch {
            $findings += New-PMFinding -Severity 'WARN' -TextKey 'agsusage.finding.dataError' -Values @($reportName, $_.Exception.Message)
            return New-PMResult -Id 'AGSUSAGE' -TitleKey 'agsusage.title' -Status 'WARN' `
                -SummaryKey 'agsusage.summary.error' -SummaryValues @($reportName) -Findings $findings
        }

        $report = $resp.report
        if (-not $report) {
            $findings += New-PMFinding -Severity 'WARN' -TextKey 'agsusage.finding.dataError' -Values @($reportName, 'the response carried no report data')
            return New-PMResult -Id 'AGSUSAGE' -TitleKey 'agsusage.title' -Status 'WARN' `
                -SummaryKey 'agsusage.summary.error' -SummaryValues @($reportName) -Findings $findings
        }

        # "time-slices" and "report-data" are the literal JSON property names
        # and are not valid bare PowerShell identifiers (the hyphen), hence
        # the quoted member access below rather than dot notation.
        $metrics     = @($report.metadata.metrics)
        $reportData  = @($report.'report-data')

        # report-data[i] lines up positionally with metadata.metrics[i]; each
        # entry there is itself a list of resource entries (normally one, for
        # the site-wide "services/*" resource) carrying a "data" array
        # aligned with time-slices. Every level is read defensively because
        # this shape has never been confirmed against a real site response -
        # see the "unverified" note in HANDOVER.md.
        $values = @{}
        for ($i = 0; $i -lt $metrics.Count; $i++) {
            $metricName = [string]$metrics[$i].metric
            if ([string]::IsNullOrWhiteSpace($metricName)) { continue }

            $nums = @()
            if ($i -lt $reportData.Count) {
                foreach ($entry in @($reportData[$i])) {
                    foreach ($v in @($entry.data)) {
                        if ($null -ne $v) { $nums += [double]$v }
                    }
                }
            }
            if ($nums.Count -eq 0) { continue }

            # A metric named "...Max..." is a peak and must not be summed
            # across time slices - everything else here is a count, and
            # counts are summed.
            if ($metricName -match 'Max') { $values[$metricName] = ($nums | Measure-Object -Maximum).Maximum }
            else                          { $values[$metricName] = ($nums | Measure-Object -Sum).Sum }
        }

        $columns = New-PMItemColumns
        $rows    = @()

        $rows += New-PMItemRow -TextKey 'agsusage.item.report' -Value $reportName

        $period = ''
        $start = $report.metadata.temporalinfo.startTime
        $end   = $report.metadata.temporalinfo.endTime
        if ($start -and $end -and [double]$start -gt 0 -and [double]$end -gt 0) {
            $epoch = New-Object DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
            $sDate = $epoch.AddMilliseconds([double]$start).ToLocalTime()
            $eDate = $epoch.AddMilliseconds([double]$end).ToLocalTime()
            $period = "{0:yyyy-MM-dd HH:mm} - {1:yyyy-MM-dd HH:mm}" -f $sDate, $eDate
        }
        elseif ($report.since) { $period = [string]$report.since }
        if ($period) { $rows += New-PMItemRow -TextKey 'agsusage.item.period' -Value $period }

        $timedOut = $null
        if ($values.ContainsKey('RequestsTimedOut')) {
            $timedOut = $values['RequestsTimedOut']
            $rows += New-PMItemRow -TextKey 'agsusage.item.timedOut' -Value $timedOut
        }
        if ($values.ContainsKey('RequestCount')) {
            $rows += New-PMItemRow -TextKey 'agsusage.item.totalRequests' -Value $values['RequestCount']
        }
        if ($values.ContainsKey('RequestMaxResponseTime')) {
            $rows += New-PMItemRow -TextKey 'agsusage.item.maxResponseMs' -Value $values['RequestMaxResponseTime']
        }

        # Any metric this site's report happens to carry beyond the three
        # named above is still shown, generically, rather than dropped -
        # usage reports are configurable per site and the set queried here
        # is only what one real site was seen to have enabled.
        foreach ($key in $values.Keys) {
            if ($key -in @('RequestsTimedOut', 'RequestCount', 'RequestMaxResponseTime')) { continue }
            $rows += @{ Item = $key; ItemEn = $key; Value = $values[$key]; ValueEn = ''; _RowStatus = '' }
        }

        if ($values.Count -eq 0) {
            $findings += New-PMFinding -Severity 'INFO' -TextKey 'agsusage.finding.noMetrics' -Values @($reportName)
        }

        if ($null -ne $timedOut -and $timedOut -gt 0) {
            $status = Test-PMThreshold -Name 'AGSUsageTimedOutRequests' -Value $timedOut
            $findings += New-PMFinding -Severity $status -TextKey 'agsusage.finding.timedOut' -Values @($timedOut, $reportName)
            $sumKey = 'agsusage.summary.issue'
            $sumVal = @($reportName, $timedOut)
        }
        else {
            $status = 'OK'
            $sumKey = 'agsusage.summary.ok'
            $sumVal = @($reportName)
        }

        return New-PMResult -Id 'AGSUSAGE' -TitleKey 'agsusage.title' -Status $status `
            -SummaryKey $sumKey -SummaryValues $sumVal `
            -Columns $columns -Rows $rows -Findings $findings `
            -Raw ([pscustomobject]@{ ReportName = $reportName; AvailableReports = $names; Metrics = $values })
    }
    finally {
        Restore-PMArcGISCertificatePolicy
    }
}

Register-PMCheck -Id 'AGSUSAGE' -TitleKey 'agsusage.title' -Function 'Invoke-PMCheckArcGISUsage'
