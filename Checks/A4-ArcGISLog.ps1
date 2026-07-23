# PMtools check - ArcGIS Server log messages at SEVERE and WARNING level,
# grouped and counted the same way Checks\50-EventLog.ps1 groups Windows
# Event Log entries. ASCII-only; text comes from i18n.json.
#
# Needs a configured connection and is disabled by default, like every other
# A*-ArcGIS* check - see A0-ArcGISSite.ps1.
#
# POST /admin/logs/query, because the Admin API requires it here (the filter
# is a JSON object too large for a query string), not because it writes -
# this only reads the log the site already keeps.
#
# Schema confirmed against a real site (ArcGIS Server 11.5.0) on 2026-07-23:
#
#  - The "filter" parameter is REQUIRED, same trap as AGSUSAGE's report
#    query - omitting it answers "Invalid filter value. / missing query
#    parameters" rather than falling back to no filtering. An empty-ish
#    filter ({"codes":[],"processIds":[],"users":[],"server":"*"}) is
#    accepted and matches everything.
#
#  - "level" is a THRESHOLD, not an exact match: level=WARNING returns
#    WARNING messages *and* SEVERE ones mixed together, newest first.
#    Confirmed by querying level=SEVERE alone and getting a different,
#    smaller result. This matters because a single frequent WARNING can
#    fill an entire page and push genuine SEVERE entries out of it - a
#    real site hit exactly this (one noisy "service not found" WARNING
#    repeated dozens of times in the most recent entries, while a real
#    SEVERE error sat further back, invisible to a WARNING-only query).
#    Queried separately per level below for that reason, and each
#    result set is filtered to its own exact "type" afterwards so a
#    SEVERE message pulled in by the WARNING query is not double-counted.
#
#  - Response shape is {hasMore, startTime, endTime, logMessages: [...]},
#    not the {hasMultipleErrorReported, ...} shape assumed before looking.
#    Each message carries type, message, time (epoch ms), source, machine,
#    user, code, process, thread, requestID.
#
# pageSize is capped (ArcGIS.LogPageSize, default 100) deliberately: a busy
# site's log can run into the tens of thousands of rows for a 7-day window,
# and this check exists to flag that something is wrong, not to be a log
# viewer. "hasMore" is carried into the result so a capped count is never
# mistaken for a complete one.

function Invoke-PMCheckArcGISLog {

    try {
        $session   = Get-PMArcGISSession
        $days      = [int](Get-PMSetting -Path 'ArcGIS.LogLookbackDays' -Default 7)
        $pageSize  = [int](Get-PMSetting -Path 'ArcGIS.LogPageSize'     -Default 100)
        $maxGroups = [int](Get-PMSetting -Path 'ArcGIS.LogMaxGroups'    -Default 25)

        $epoch     = New-Object DateTime -ArgumentList 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
        $endTime   = [int64](((Get-Date).ToUniversalTime()) - $epoch).TotalMilliseconds
        $startTime = [int64](((Get-Date).ToUniversalTime().AddDays(-$days)) - $epoch).TotalMilliseconds
        $filter    = [pscustomobject]@{ codes = @(); processIds = @(); users = @(); server = '*' } |
                     ConvertTo-Json -Compress

        function Get-PMArcGISLogLevel {
            param([string]$Level)
            return Invoke-PMArcGISAdmin -Root $session.Root -Path 'logs/query' -Token $session.Token `
                                         -TimeoutSec $session.TimeoutSec -Method Post -Parameters @{
                startTime = $startTime; endTime = $endTime; level = $Level; pageSize = $pageSize; filter = $filter
            }
        }

        $findings   = @()
        $allGroups  = @()
        $hasMore    = $false
        $anySuccess = $false
        $counts     = @{ SEVERE = 0; WARNING = 0 }

        foreach ($level in @('SEVERE', 'WARNING')) {
            $resp    = $null
            $errText = ''
            try {
                $resp = Get-PMArcGISLogLevel -Level $level
            }
            catch { $errText = $_.Exception.Message }

            # logs/query reports its own failures as {"status":"error",
            # "messages":[...]} in an HTTP 200 body - confirmed against a
            # real site (the very first request here, sent without the
            # required "filter" parameter, came back this way). It does
            # NOT use the {"error":{...}} shape Invoke-PMArcGISAdmin's
            # generic handling throws on, so that check alone is not
            # enough: a permissions error here would otherwise come back
            # as $resp.logMessages = $null, which the code below would
            # silently read as "zero problems found" - a false OK on a
            # query that never actually ran. Checked explicitly instead.
            if ([string]::IsNullOrWhiteSpace($errText) -and $resp -and $resp.PSObject.Properties['status'] -and [string]$resp.status -eq 'error') {
                $errText = if ($resp.messages) { @($resp.messages) -join '; ' } else { 'unknown error' }
                $resp = $null
            }

            if ([string]::IsNullOrWhiteSpace($errText) -and $null -eq $resp) { $errText = 'no response' }

            if (-not [string]::IsNullOrWhiteSpace($errText)) {
                $findings += New-PMFinding -Severity 'WARN' -TextKey 'agslog.finding.queryError' -Values @($level, $errText)
                continue
            }

            $anySuccess = $true
            if ($resp.hasMore) { $hasMore = $true }

            # WARNING-level query also carries SEVERE messages ahead of it in
            # the same page (level is a threshold - see header comment), so
            # only entries whose own type matches this pass are counted here.
            # SEVERE entries are already captured in full by the SEVERE pass.
            $msgs = @($resp.logMessages | Where-Object { [string]$_.type -eq $level })
            $counts[$level] = $msgs.Count

            $allGroups += @($msgs |
                Group-Object -Property { "{0}|{1}|{2}" -f $_.machine, $_.code, $_.type } |
                ForEach-Object {
                    $first = $_.Group[0]
                    $last  = ($_.Group | Measure-Object -Property time -Maximum).Maximum
                    [pscustomobject]@{
                        Level    = [string]$first.type
                        Machine  = [string]$first.machine
                        Code     = [string]$first.code
                        Count    = $_.Count
                        LastTime = $epoch.AddMilliseconds([double]$last).ToLocalTime()
                        Sample   = Get-PMShortText -Text $first.message -MaxLength 180
                    }
                })
        }

        $severeCount  = $counts['SEVERE']
        $warningCount = $counts['WARNING']
        $totalCount   = $severeCount + $warningCount

        $columns = @(
            (New-PMColumn -Key 'Level'   -TextKey 'agslog.col.level'),
            (New-PMColumn -Key 'Machine' -TextKey 'agslog.col.machine'),
            (New-PMColumn -Key 'Code'    -TextKey 'agslog.col.code' -Align 'right'),
            (New-PMColumn -Key 'Count'   -TextKey 'agslog.col.count' -Align 'right'),
            (New-PMColumn -Key 'Last'    -TextKey 'agslog.col.last'),
            (New-PMColumn -Key 'Sample'  -TextKey 'agslog.col.sample' -Wide)
        )

        $severeWord  = Get-PMWord -Key 'agslog.level.severe'
        $warningWord = Get-PMWord -Key 'agslog.level.warning'

        $sortedGroups = @($allGroups | Sort-Object @{ Expression = { if ($_.Level -eq 'SEVERE') { 0 } else { 1 } } }, @{ Expression = 'Count'; Descending = $true })
        $shown        = @($sortedGroups | Select-Object -First $maxGroups)

        $rows = @()
        foreach ($g in $shown) {
            if ($g.Level -eq 'SEVERE') { $status = 'CRIT'; $lvTh = $severeWord.Th; $lvEn = $severeWord.En }
            else                       { $status = 'WARN'; $lvTh = $warningWord.Th; $lvEn = $warningWord.En }

            $row = New-PMRow -Status $status -Values @{
                Level   = $lvTh
                Machine = $g.Machine
                Code    = $g.Code
                Count   = $g.Count
                Last    = Format-PMStamp $g.LastTime
                Sample  = $g.Sample
            }
            $row['LevelEn'] = $lvEn
            $rows += $row
        }

        foreach ($g in @($sortedGroups | Where-Object { $_.Level -eq 'SEVERE' } | Select-Object -First 5)) {
            $findings += New-PMFinding -Severity 'CRIT' -TextKey 'agslog.finding.severe' `
                -Values @($g.Machine, $g.Code, $g.Count, (Format-PMStamp $g.LastTime))
        }
        foreach ($g in @($sortedGroups | Where-Object { $_.Level -eq 'WARNING' } | Select-Object -First 3)) {
            $findings += New-PMFinding -Severity 'WARN' -TextKey 'agslog.finding.warning' `
                -Values @($g.Machine, $g.Code, $g.Count, (Format-PMStamp $g.LastTime))
        }
        if ($hasMore) {
            $findings += New-PMFinding -Severity 'INFO' -TextKey 'agslog.finding.capped' -Values @($pageSize)
        }

        if (-not $anySuccess) {
            # Neither query came back - "zero problems found" would be
            # indistinguishable from a genuinely healthy site, and this is
            # not that. Say so instead of reporting OK on data never seen.
            $status = 'WARN'
            $sumKey = 'agslog.summary.queryError'
            $sumVal = @($days)
        }
        elseif ($totalCount -eq 0) {
            $status = 'OK'
            $sumKey = 'agslog.summary.ok'
            $sumVal = @($days)
        }
        else {
            if ($severeCount -gt 0) { $status = 'CRIT' } else { $status = 'WARN' }
            $sumKey = 'agslog.summary.issue'
            $sumVal = @($severeCount, $warningCount, $days)
        }

        return New-PMResult -Id 'AGSLOG' -TitleKey 'agslog.title' -Status $status `
            -SummaryKey $sumKey -SummaryValues $sumVal `
            -Columns $columns -Rows $rows -Findings $findings `
            -Raw ([pscustomobject]@{
                LookbackDays = $days
                PageSize     = $pageSize
                HasMore      = $hasMore
                SevereCount  = $severeCount
                WarningCount = $warningCount
                GroupCount   = $allGroups.Count
                Groups       = $shown
            })
    }
    finally {
        Restore-PMArcGISCertificatePolicy
    }
}

Register-PMCheck -Id 'AGSLOG' -TitleKey 'agslog.title' -Function 'Invoke-PMCheckArcGISLog'
