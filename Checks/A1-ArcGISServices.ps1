# PMtools check - ArcGIS Server service inventory, state and load.
# ASCII-only; text comes from i18n.json.
#
# Needs a configured connection and is disabled by default, like every other
# A*-ArcGIS* check - see A0-ArcGISSite.ps1.
#
# Read-only: /services and /services/<folder>/report are both plain GETs.
#
# Deliberately NOT using usage reports for the request numbers. Querying one
# needs /usagereports/<name>/data, which only works if a report already
# exists on the site, and creating one is a write. The per-service
# "transactions" and "totalBusyTime" counters below come back from the plain
# services report instead, and answer the same operational question.

function Invoke-PMCheckArcGISServices {

    try {
        $session = Get-PMArcGISSession
        $topCount = [int](Get-PMSetting -Path 'ArcGIS.TopServiceCount' -Default 10)

        # The root report covers only services sitting outside any folder, so
        # every folder has to be asked separately. '' is the root.
        $folders = @('')
        $root = Invoke-PMArcGISAdmin -Root $session.Root -Path 'services' -Token $session.Token -TimeoutSec $session.TimeoutSec
        if ($root.folders) { $folders += @($root.folders) }

        $services = @()
        $findings = @()

        foreach ($folder in $folders) {
            if ([string]::IsNullOrWhiteSpace($folder)) { $path = 'services/report' }
            else                                       { $path = "services/$folder/report" }

            try {
                $rep = Invoke-PMArcGISAdmin -Root $session.Root -Path $path -Token $session.Token -TimeoutSec $session.TimeoutSec
            }
            catch {
                $label = if ($folder) { $folder } else { '/' }
                $findings += New-PMFinding -Severity 'WARN' -TextKey 'agssvc.finding.folderError' -Values @($label, $_.Exception.Message)
                continue
            }

            foreach ($r in @($rep.reports)) {
                $inst = $r.instances
                $max  = 0; $busy = 0; $tx = 0; $busyTime = 0
                if ($inst) {
                    if ($inst.max)           { $max      = [int]$inst.max }
                    if ($inst.busy)          { $busy     = [int]$inst.busy }
                    if ($inst.transactions)  { $tx       = [long]$inst.transactions }
                    if ($inst.totalBusyTime) { $busyTime = [long]$inst.totalBusyTime }
                }

                $folderName = [string]$r.folderName
                if ([string]::IsNullOrWhiteSpace($folderName)) { $folderName = '/' }

                # totalBusyTime is the server-side milliseconds those
                # transactions spent occupying an instance, so the quotient is
                # average processing time - NOT round-trip response time,
                # which would also include queueing and the network. Labelled
                # as processing time in i18n for that reason.
                if ($tx -gt 0) { $avgMs = [math]::Round(($busyTime / $tx), 1) } else { $avgMs = $null }

                $services += [pscustomobject]@{
                    Folder     = $folderName
                    Name       = [string]$r.serviceName
                    Type       = [string]$r.type
                    Configured = [string]$r.status.configuredState
                    RealTime   = [string]$r.status.realTimeState
                    Max        = $max
                    Busy       = $busy
                    Requests   = $tx
                    BusyTimeMs = $busyTime
                    AvgMs      = $avgMs
                }
            }
        }

        if ($services.Count -eq 0) {
            return New-PMResult -Id 'AGSSVC' -TitleKey 'agssvc.title' -Status 'INFO' `
                -SummaryKey 'agssvc.summary.none' -Findings $findings
        }

        # Classify. The distinction that matters is configured vs real, NOT
        # "is it running": a site ships with a dozen System and Utilities
        # services deliberately stopped, and on the first real site tested
        # sixteen of them were. Treating "not started" as a fault would have
        # raised sixteen critical alarms on a perfectly healthy production
        # site - the same mistake already made once with machine state, at
        # larger scale. Only configured=STARTED but real<>STARTED is a fault.
        $running = 0; $failed = 0; $byDesign = 0; $unexpected = 0; $saturated = 0

        foreach ($svc in $services) {
            $cfgStarted  = ($svc.Configured -eq 'STARTED')
            $realStarted = ($svc.RealTime -eq 'STARTED')
            $label = "$($svc.Folder)/$($svc.Name)"

            if ($cfgStarted -and $realStarted) {
                $running++
                if ($svc.Max -gt 0 -and $svc.Busy -ge $svc.Max) {
                    $saturated++
                    $svc | Add-Member -NotePropertyName Status -NotePropertyValue 'WARN' -Force
                    $svc | Add-Member -NotePropertyName StateKey -NotePropertyValue 'agssvc.state.saturated' -Force
                    $findings += New-PMFinding -Severity 'WARN' -TextKey 'agssvc.finding.saturated' -Values @($label, $svc.Max)
                }
                else {
                    $svc | Add-Member -NotePropertyName Status -NotePropertyValue 'OK' -Force
                    $svc | Add-Member -NotePropertyName StateKey -NotePropertyValue 'agssvc.state.started' -Force
                }
            }
            elseif ($cfgStarted -and -not $realStarted) {
                $failed++
                $svc | Add-Member -NotePropertyName Status -NotePropertyValue 'CRIT' -Force
                $svc | Add-Member -NotePropertyName StateKey -NotePropertyValue 'agssvc.state.failed' -Force
                $findings += New-PMFinding -Severity 'CRIT' -TextKey 'agssvc.finding.failed' -Values @($label, $svc.Type)
            }
            elseif (-not $cfgStarted -and $realStarted) {
                $unexpected++
                $svc | Add-Member -NotePropertyName Status -NotePropertyValue 'WARN' -Force
                $svc | Add-Member -NotePropertyName StateKey -NotePropertyValue 'agssvc.state.unexpected' -Force
                $findings += New-PMFinding -Severity 'WARN' -TextKey 'agssvc.finding.unexpected' -Values @($label)
            }
            else {
                $byDesign++
                $svc | Add-Member -NotePropertyName Status -NotePropertyValue 'INFO' -Force
                $svc | Add-Member -NotePropertyName StateKey -NotePropertyValue 'agssvc.state.stopped' -Force
            }
        }

        # A site can carry well over a hundred services. Show everything that
        # needs attention, then fill up with the busiest - the same shape as
        # the EVT and PROC checks, which cap their tables the same way. The
        # full list always reaches PM-Data.json regardless.
        $attention = @($services | Where-Object { $_.Status -eq 'CRIT' -or $_.Status -eq 'WARN' })
        $rest      = @($services | Where-Object { $_.Status -ne 'CRIT' -and $_.Status -ne 'WARN' -and $_.Requests -gt 0 } |
                       Sort-Object Requests -Descending | Select-Object -First $topCount)
        $shown     = @($attention) + @($rest)

        $columns = @(
            (New-PMColumn -Key 'Service'   -TextKey 'agssvc.col.service' -Wide),
            (New-PMColumn -Key 'Type'      -TextKey 'agssvc.col.type'),
            (New-PMColumn -Key 'Instances' -TextKey 'agssvc.col.instances' -Align 'right'),
            (New-PMColumn -Key 'Requests'  -TextKey 'agssvc.col.requests'  -Align 'right'),
            (New-PMColumn -Key 'AvgMs'     -TextKey 'agssvc.col.avgms'     -Align 'right')
        )

        $rows = @()
        foreach ($svc in $shown) {
            $state = Get-PMWord -Key $svc.StateKey
            if ($svc.Max -gt 0) { $instances = "$($svc.Busy) / $($svc.Max)" } else { $instances = '-' }
            if ($null -eq $svc.AvgMs) { $avg = '-' } else { $avg = $svc.AvgMs }

            $rows += @{
                Service    = "$($svc.Folder)/$($svc.Name)"
                Type       = $svc.Type
                Instances  = $instances
                Requests   = $svc.Requests
                AvgMs      = $avg
                State      = $state.Th
                StateEn    = $state.En
                _RowStatus = $svc.Status
            }
        }

        if ($shown.Count -lt $services.Count) {
            $note = Get-PMText -Key 'agssvc.note.limited' -Values @($shown.Count, $services.Count)
            $findings += New-PMFinding -Severity 'INFO' -TextKey 'agssvc.note.limited' -Values @($shown.Count, $services.Count)
        }

        if ($failed -gt 0) {
            $status = 'CRIT'
            $sumKey = 'agssvc.summary.issue'
            $sumVal = @($services.Count, $failed)
        }
        else {
            $status = Get-PMWorstStatus (@($services | ForEach-Object { $_.Status }) + @('OK'))
            $sumKey = 'agssvc.summary.ok'
            $sumVal = @($services.Count, $running, $byDesign)
        }

        return New-PMResult -Id 'AGSSVC' -TitleKey 'agssvc.title' -Status $status `
            -SummaryKey $sumKey -SummaryValues $sumVal `
            -Columns $columns -Rows $rows -Findings $findings `
            -Raw ([pscustomobject]@{
                Total = $services.Count; Running = $running; FailedToStart = $failed
                StoppedByDesign = $byDesign; RunningUnexpectedly = $unexpected; Saturated = $saturated
                Services = $services
            })
    }
    finally {
        Restore-PMArcGISCertificatePolicy
    }
}

Register-PMCheck -Id 'AGSSVC' -TitleKey 'agssvc.title' -Function 'Invoke-PMCheckArcGISServices'
