# PMtools check - ArcGIS Server usage report totals (requests, response time,
# timeouts) over whatever window the site's own usage reports already cover.
# ASCII-only; text comes from i18n.json.
#
# Needs a configured connection and is disabled by default, like every other
# A*-ArcGIS* check - see A0-ArcGISSite.ps1.
#
# Read-only, and stricter than the other ArcGIS checks about it: this check
# is ONLY allowed to read usage reports that already exist on the site.
# /usagereports/<name>/data can only answer for a report that has already
# been created - and creating one is a write.
#
# Schema confirmed against a real site (ArcGIS Server 11.5.0) on 2026-07-23,
# and it differs from Esri's general documentation in ways that matter:
#
#  - /admin/usagereports returns its list under a "metrics" property, not
#    "usagereports". Each entry is a REPORT DEFINITION (reportname, a
#    "since" window label, one or more "queries" of resourceURIs+metrics,
#    and metadata.temp) - not report data.
#
#  - ArcGIS Server Manager creates a report every time someone opens its
#    Statistics page, with metadata.temp = true and a reportname that is a
#    raw millisecond timestamp. Those are excluded here on purpose: the
#    name is not stable across runs and the report can expire on its own
#    (its metadata carries a "tempTimer"). Only metadata.temp = false
#    (permanent) reports are read - on the test site there were three,
#    each covering exactly one metric: RequestCount, RequestMaxResponseTime
#    and RequestsTimedOut, one metric per report rather than one report
#    with three.
#
#  - Querying /usagereports/<name>/data with no other parameters does NOT
#    fall back to the report's own stored query - it fails server-side with
#    a bare HTTP 500 ("RequestUtil.getParameterIgnoreCase(...) is null").
#    A "filter" parameter is required on every call, and the report's own
#    "since" and "queries" (both already in hand from the list call above)
#    are exactly what it wants, JSON-encoded.
#
#  - Each report-data entry is self-labelled with "metric-type", so metrics
#    are grouped by that label directly. There is no need to line anything
#    up positionally against metadata, which is fortunate, because the
#    metadata field comes back as a JSON-encoded STRING inside the data
#    response (unlike the list response, where it is a real object).

function Invoke-PMCheckArcGISUsage {

    try {
        $session = Get-PMArcGISSession

        $list = Invoke-PMArcGISAdmin -Root $session.Root -Path 'usagereports' -Token $session.Token -TimeoutSec $session.TimeoutSec

        $reports = @(@($list.metrics) | Where-Object {
            $_ -and $_.metadata -and $_.metadata.PSObject.Properties['temp'] -and $_.metadata.temp -eq $false -and $_.queries
        })

        if ($reports.Count -eq 0) {
            return New-PMResult -Id 'AGSUSAGE' -TitleKey 'agsusage.title' -Status 'INFO' `
                -SummaryKey 'agsusage.summary.none'
        }

        $findings   = @()
        $values     = @{}
        $usedNames  = @()
        $failedNames = @()
        $allSlices  = @()

        foreach ($rep in $reports) {
            $reportName = [string]$rep.reportname
            $filter     = [pscustomobject]@{ since = $rep.since; queries = $rep.queries } | ConvertTo-Json -Depth 6 -Compress
            $encName    = [Uri]::EscapeDataString($reportName)

            try {
                $resp = Invoke-PMArcGISAdmin -Root $session.Root -Path "usagereports/$encName/data" `
                                              -Token $session.Token -TimeoutSec $session.TimeoutSec `
                                              -Parameters @{ filter = $filter }
            }
            catch {
                $failedNames += $reportName
                $findings += New-PMFinding -Severity 'WARN' -TextKey 'agsusage.finding.dataError' -Values @($reportName, $_.Exception.Message)
                continue
            }

            $report = $resp.report
            if (-not $report) {
                $failedNames += $reportName
                $findings += New-PMFinding -Severity 'WARN' -TextKey 'agsusage.finding.dataError' -Values @($reportName, 'the response carried no report data')
                continue
            }

            $usedNames += $reportName
            foreach ($slice in @($report.'time-slices')) { if ($null -ne $slice) { $allSlices += [double]$slice } }

            foreach ($group in @($report.'report-data')) {
                foreach ($entry in @($group)) {
                    $metricName = [string]$entry.'metric-type'
                    if ([string]::IsNullOrWhiteSpace($metricName)) { continue }

                    $nums = @()
                    foreach ($v in @($entry.data)) { if ($null -ne $v) { $nums += [double]$v } }
                    if ($nums.Count -eq 0) { continue }

                    # A metric named "...Max..." is a peak and must not be
                    # summed across time slices or across reports -
                    # everything else here is a count, and counts are
                    # summed both within a report and across reports (in
                    # case a site ever splits the same metric across more
                    # than one permanent report).
                    if ($metricName -match 'Max') {
                        $agg = ($nums | Measure-Object -Maximum).Maximum
                        if ($values.ContainsKey($metricName)) { $values[$metricName] = [math]::Max($values[$metricName], $agg) }
                        else                                  { $values[$metricName] = $agg }
                    }
                    else {
                        $agg = ($nums | Measure-Object -Sum).Sum
                        if ($values.ContainsKey($metricName)) { $values[$metricName] += $agg }
                        else                                  { $values[$metricName] = $agg }
                    }
                }
            }
        }

        if ($usedNames.Count -eq 0) {
            return New-PMResult -Id 'AGSUSAGE' -TitleKey 'agsusage.title' -Status 'WARN' `
                -SummaryKey 'agsusage.summary.error' -SummaryValues @(($failedNames -join ', ')) -Findings $findings
        }

        if ($values.Count -eq 0) {
            $findings += New-PMFinding -Severity 'INFO' -TextKey 'agsusage.finding.noMetrics' -Values @(($usedNames -join ', '))
        }

        $columns = New-PMItemColumns
        $rows    = @()

        $rows += New-PMItemRow -TextKey 'agsusage.item.report' -Value ($usedNames -join ', ')

        if ($allSlices.Count -gt 0) {
            $epoch = New-Object DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
            $sDate = $epoch.AddMilliseconds(($allSlices | Measure-Object -Minimum).Minimum).ToLocalTime()
            $eDate = $epoch.AddMilliseconds(($allSlices | Measure-Object -Maximum).Maximum).ToLocalTime()
            $period = "{0:yyyy-MM-dd HH:mm} - {1:yyyy-MM-dd HH:mm}" -f $sDate, $eDate
            $rows += New-PMItemRow -TextKey 'agsusage.item.period' -Value $period
        }

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

        # Any metric a site's permanent reports happen to carry beyond the
        # three named above is still shown, generically, rather than
        # dropped - which metrics exist is a per-site Manager configuration
        # choice, not something this check controls.
        foreach ($key in $values.Keys) {
            if ($key -in @('RequestsTimedOut', 'RequestCount', 'RequestMaxResponseTime')) { continue }
            $rows += @{ Item = $key; ItemEn = $key; Value = $values[$key]; ValueEn = ''; _RowStatus = '' }
        }

        if ($failedNames.Count -gt 0) {
            $findings += New-PMFinding -Severity 'INFO' -TextKey 'agsusage.note.partial' -Values @(($failedNames -join ', '))
        }

        if ($null -ne $timedOut -and $timedOut -gt 0) {
            $status = Test-PMThreshold -Name 'AGSUsageTimedOutRequests' -Value $timedOut
            $findings += New-PMFinding -Severity $status -TextKey 'agsusage.finding.timedOut' -Values @($timedOut, ($usedNames -join ', '))
            $sumKey = 'agsusage.summary.issue'
            $sumVal = @(($usedNames -join ', '), $timedOut)
        }
        else {
            $status = 'OK'
            $sumKey = 'agsusage.summary.ok'
            $sumVal = @(($usedNames -join ', '))
        }

        return New-PMResult -Id 'AGSUSAGE' -TitleKey 'agsusage.title' -Status $status `
            -SummaryKey $sumKey -SummaryValues $sumVal `
            -Columns $columns -Rows $rows -Findings $findings `
            -Raw ([pscustomobject]@{ ReportsUsed = $usedNames; ReportsFailed = $failedNames; Metrics = $values })
    }
    finally {
        Restore-PMArcGISCertificatePolicy
    }
}

Register-PMCheck -Id 'AGSUSAGE' -TitleKey 'agsusage.title' -Function 'Invoke-PMCheckArcGISUsage'
