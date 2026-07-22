# PMtools check - Service state.  ASCII-only; text comes from i18n.json.
#
# Two independent signals:
#   1. any service set to start automatically that is not running  -> WARN
#   2. any service named in Config\services.json that is not running -> CRIT
# Signal 1 needs no configuration at all, so the check is useful on a server
# nobody has tuned yet; signal 2 is how a site marks what actually matters.

function Invoke-PMCheckServices {

    $all       = @(Get-CimInstance -ClassName Win32_Service -ErrorAction Stop)
    $watchlist = @(Get-PMWatchlist)

    $autoServices = @($all | Where-Object { $_.StartMode -eq 'Auto' })

    $columns = @(
        (New-PMColumn -Key 'Name'      -TextKey 'svc.col.name'),
        (New-PMColumn -Key 'Display'   -TextKey 'svc.col.display' -Wide),
        (New-PMColumn -Key 'StartMode' -TextKey 'svc.col.startmode'),
        (New-PMColumn -Key 'State'     -TextKey 'svc.col.state'),
        (New-PMColumn -Key 'Watch'     -TextKey 'svc.col.watch' -Align 'center')
    )

    $yes = Get-PMWord -Key 'common.yes'
    $no  = Get-PMWord -Key 'common.no'

    # Show every watchlisted service whatever its state (so the reader can see
    # the important ones were actually looked at), plus any auto-start service
    # that is stopped.
    $interesting = @($all | Where-Object {
        ($watchlist -contains $_.Name) -or
        ($_.StartMode -eq 'Auto' -and $_.State -ne 'Running')
    })

    $rows     = @()
    $findings = @()
    $raw      = @()
    $stopped  = 0

    foreach ($s in ($interesting | Sort-Object Name)) {

        $isWatched = ($watchlist -contains $s.Name)
        $isRunning = ($s.State -eq 'Running')

        if ($isRunning)       { $status = 'OK' }
        elseif ($isWatched)   { $status = 'CRIT'; $stopped++ }
        else                  { $status = 'WARN'; $stopped++ }

        if ($isWatched) { $watchTh = $yes.Th; $watchEn = $yes.En } else { $watchTh = $no.Th; $watchEn = $no.En }

        $row = New-PMRow -Status $status -Values @{
            Name      = $s.Name
            Display   = $s.DisplayName
            StartMode = $s.StartMode
            State     = $s.State
            Watch     = $watchTh
        }
        $row['WatchEn'] = $watchEn
        $rows += $row

        if (-not $isRunning) {
            if ($isWatched) { $findings += New-PMFinding -Severity 'CRIT' -TextKey 'svc.finding.watch' -Values @($s.DisplayName, $s.Name) }
            else            { $findings += New-PMFinding -Severity 'WARN' -TextKey 'svc.finding.auto'  -Values @($s.DisplayName, $s.Name) }
        }

        $raw += [pscustomobject]@{
            Name = $s.Name; DisplayName = $s.DisplayName; StartMode = $s.StartMode
            State = $s.State; Watchlisted = $isWatched; Status = $status
        }
    }

    $overall = Get-PMWorstStatus (@($rows | ForEach-Object { $_._RowStatus }) + @('OK'))

    if ($stopped -gt 0) { $sumKey = 'svc.summary.issue'; $sumVal = @($stopped, $autoServices.Count) }
    else                { $sumKey = 'svc.summary.ok';    $sumVal = @($autoServices.Count) }

    return New-PMResult -Id 'SVC' -TitleKey 'svc.title' -Status $overall `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns $columns -Rows $rows -Findings $findings `
        -Raw ([pscustomobject]@{
            TotalServices     = $all.Count
            AutoStartServices = $autoServices.Count
            WatchlistCount    = $watchlist.Count
            Reported          = $raw
        })
}

Register-PMCheck -Id 'SVC' -TitleKey 'svc.title' -Function 'Invoke-PMCheckServices'
