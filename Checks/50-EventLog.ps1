# PMtools check - Critical and Error events.  ASCII-only; text comes from i18n.json.
#
# Events are grouped by log + level + provider + event id and reported as counts.
# A raw dump of several hundred near-identical events is unreadable in a report
# and hides the one entry that matters.
#
# Grouping uses the numeric Level (1 = Critical, 2 = Error) rather than
# LevelDisplayName, which is localised and would group differently per language.

function Invoke-PMCheckEventLog {

    $days      = [int](Get-PMSetting -Path 'Lookback.EventLogDays'      -Default 7)
    $maxEvents = [int](Get-PMSetting -Path 'Lookback.EventLogMaxEvents' -Default 3000)
    $maxGroups = [int](Get-PMSetting -Path 'Lookback.EventLogMaxGroups' -Default 25)
    $since     = (Get-Date).AddDays(-$days)

    $events    = @()
    $truncated = $false

    foreach ($logName in @('System', 'Application')) {
        $found = @(Get-WinEvent -FilterHashtable @{ LogName = $logName; Level = 1, 2; StartTime = $since } `
                                -MaxEvents $maxEvents -ErrorAction SilentlyContinue)
        if ($found.Count -ge $maxEvents) { $truncated = $true }
        $events += $found
    }

    $columns = @(
        (New-PMColumn -Key 'Log'      -TextKey 'evt.col.log'),
        (New-PMColumn -Key 'Level'    -TextKey 'evt.col.level'),
        (New-PMColumn -Key 'Provider' -TextKey 'evt.col.provider'),
        (New-PMColumn -Key 'EventId'  -TextKey 'evt.col.id'    -Align 'right'),
        (New-PMColumn -Key 'Count'    -TextKey 'evt.col.count' -Align 'right'),
        (New-PMColumn -Key 'Last'     -TextKey 'evt.col.last'),
        (New-PMColumn -Key 'Sample'   -TextKey 'evt.col.sample' -Wide)
    )

    $critWord  = Get-PMWord -Key 'evt.level.crit'
    $errorWord = Get-PMWord -Key 'evt.level.error'

    $groups = @($events |
        Group-Object -Property { "{0}|{1}|{2}|{3}" -f $_.LogName, $_.Level, $_.ProviderName, $_.Id } |
        ForEach-Object {
            $first = $_.Group[0]
            $last  = ($_.Group | Measure-Object -Property TimeCreated -Maximum).Maximum
            [pscustomobject]@{
                LogName  = $first.LogName
                Level    = [int]$first.Level
                Provider = $first.ProviderName
                EventId  = $first.Id
                Count    = $_.Count
                LastTime = $last
                Sample   = Get-PMShortText -Text $first.Message -MaxLength 180
            }
        } |
        Sort-Object Level, @{ Expression = 'Count'; Descending = $true })

    $critCount  = ($events | Where-Object { $_.Level -eq 1 }).Count
    $totalCount = $events.Count

    $rows = @()
    foreach ($g in ($groups | Select-Object -First $maxGroups)) {
        if ($g.Level -eq 1) { $status = 'CRIT'; $lvTh = $critWord.Th;  $lvEn = $critWord.En }
        else                { $status = 'WARN'; $lvTh = $errorWord.Th; $lvEn = $errorWord.En }

        $row = New-PMRow -Status $status -Values @{
            Log      = $g.LogName
            Level    = $lvTh
            Provider = $g.Provider
            EventId  = $g.EventId
            Count    = $g.Count
            Last     = Format-PMStamp $g.LastTime
            Sample   = $g.Sample
        }
        $row['LevelEn'] = $lvEn
        $rows += $row
    }

    # One finding per critical group, plus the three loudest error groups.
    # Beyond that the recommendations list stops being a list of actions.
    $findings = @()
    foreach ($g in @($groups | Where-Object { $_.Level -eq 1 } | Select-Object -First 5)) {
        $findings += New-PMFinding -Severity 'CRIT' -TextKey 'evt.finding.crit' `
            -Values @($g.Provider, $g.EventId, $g.Count, (Format-PMStamp $g.LastTime))
    }
    foreach ($g in @($groups | Where-Object { $_.Level -eq 2 } | Select-Object -First 3)) {
        $findings += New-PMFinding -Severity 'WARN' -TextKey 'evt.finding.error' `
            -Values @($g.Provider, $g.EventId, $g.Count, (Format-PMStamp $g.LastTime))
    }
    if ($truncated) {
        $findings += New-PMFinding -Severity 'INFO' -TextKey 'evt.finding.truncated' -Values @($maxEvents)
    }

    if ($totalCount -eq 0) {
        $status = 'OK'; $sumKey = 'evt.summary.ok'; $sumVal = @($days)
    }
    else {
        if ($critCount -gt 0) { $status = 'CRIT' } else { $status = 'WARN' }
        $sumKey = 'evt.summary.issue'; $sumVal = @($totalCount, $critCount, $days, $groups.Count)
    }

    return New-PMResult -Id 'EVT' -TitleKey 'evt.title' -Status $status `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns $columns -Rows $rows -Findings $findings `
        -Raw ([pscustomobject]@{
            LookbackDays = $days
            TotalEvents  = $totalCount
            CriticalEvents = $critCount
            GroupCount   = $groups.Count
            Truncated    = $truncated
            Groups       = @($groups | Select-Object -First $maxGroups)
        })
}

Register-PMCheck -Id 'EVT' -TitleKey 'evt.title' -Function 'Invoke-PMCheckEventLog'
