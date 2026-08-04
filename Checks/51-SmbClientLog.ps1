# PMtools check - SMB client connectivity events.
#
# Reads Microsoft-Windows-SmbClient/Connectivity, the one SMB client channel
# an admin actually watches: it is what logs disconnects, session timeouts
# and reconnect failures against mapped drives and UNC paths - the "the
# share keeps dropping" complaint. Grouped and counted the same way
# Checks\50-EventLog.ps1 groups the System/Application logs.
#
# Level 3 (Warning) is included here, unlike 50-EventLog.ps1's Critical/Error
# scope, because this channel reports most disconnect/timeout conditions at
# Warning level - a Critical/Error-only filter would miss almost everything
# it exists to surface.
#
# The channel is part of the SMB client component present on every
# supported Windows version, so - like 50-EventLog.ps1's System/Application
# logs - its presence is never separately verified: a missing channel and a
# quiet one both come back as zero events, -ErrorAction SilentlyContinue
# swallows either case the same way.
#
# No pre-flight Get-WinEvent -ListLog check on purpose: running without
# Administrator rights makes -ListLog throw "unauthorized operation" on this
# channel even though the FilterHashtable query below works fine unelevated
# - a pre-check would misreport a permissions quirk as "log not installed".

function Invoke-PMCheckSmbClientLog {

    $logName   = 'Microsoft-Windows-SmbClient/Connectivity'
    $days      = [int](Get-PMSetting -Path 'Lookback.SmbClientLogDays'      -Default 7)
    $maxEvents = [int](Get-PMSetting -Path 'Lookback.SmbClientLogMaxEvents' -Default 3000)
    $maxGroups = [int](Get-PMSetting -Path 'Lookback.SmbClientLogMaxGroups' -Default 25)
    $since     = (Get-Date).AddDays(-$days)

    $events    = @(Get-WinEvent -FilterHashtable @{ LogName = $logName; Level = 1, 2, 3; StartTime = $since } `
                                -MaxEvents $maxEvents -ErrorAction SilentlyContinue)
    $truncated = $events.Count -ge $maxEvents

    $columns = @(
        (New-PMColumn -Key 'Level'    -TextKey 'evt.col.level'),
        (New-PMColumn -Key 'Provider' -TextKey 'evt.col.provider'),
        (New-PMColumn -Key 'EventId'  -TextKey 'evt.col.id'    -Align 'right'),
        (New-PMColumn -Key 'Count'    -TextKey 'evt.col.count' -Align 'right'),
        (New-PMColumn -Key 'Last'     -TextKey 'evt.col.last'),
        (New-PMColumn -Key 'Sample'   -TextKey 'evt.col.sample' -Wide)
    )

    $critWord  = Get-PMWord -Key 'evt.level.crit'
    $errorWord = Get-PMWord -Key 'evt.level.error'
    $warnWord  = Get-PMWord -Key 'smb.level.warn'

    $groups = @($events |
        Group-Object -Property { "{0}|{1}|{2}" -f $_.Level, $_.ProviderName, $_.Id } |
        ForEach-Object {
            $first = $_.Group[0]
            $last  = ($_.Group | Measure-Object -Property TimeCreated -Maximum).Maximum
            [pscustomobject]@{
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
        elseif ($g.Level -eq 2) { $status = 'WARN'; $lvTh = $errorWord.Th; $lvEn = $errorWord.En }
        else { $status = 'WARN'; $lvTh = $warnWord.Th; $lvEn = $warnWord.En }

        $row = New-PMRow -Status $status -Values @{
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

    # One finding per critical group, plus the loudest non-critical groups -
    # same cap reasoning as 50-EventLog.ps1: past a handful, the
    # recommendations list stops being a list of actions.
    $findings = @()
    foreach ($g in @($groups | Where-Object { $_.Level -eq 1 } | Select-Object -First 5)) {
        $findings += New-PMFinding -Severity 'CRIT' -TextKey 'smb.finding.crit' `
            -Values @($g.Provider, $g.EventId, $g.Count, (Format-PMStamp $g.LastTime))
    }
    foreach ($g in @($groups | Where-Object { $_.Level -in 2, 3 } | Select-Object -First 3)) {
        $findings += New-PMFinding -Severity 'WARN' -TextKey 'smb.finding.issue' `
            -Values @($g.Provider, $g.EventId, $g.Count, (Format-PMStamp $g.LastTime))
    }
    if ($truncated) {
        $findings += New-PMFinding -Severity 'INFO' -TextKey 'smb.finding.truncated' -Values @($maxEvents)
    }

    if ($totalCount -eq 0) {
        $status = 'OK'; $sumKey = 'smb.summary.ok'; $sumVal = @($days)
    }
    else {
        if ($critCount -gt 0) { $status = 'CRIT' } else { $status = 'WARN' }
        $sumKey = 'smb.summary.issue'; $sumVal = @($totalCount, $critCount, $days, $groups.Count)
    }

    return New-PMResult -Id 'SMB' -TitleKey 'smb.title' -Status $status `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns $columns -Rows $rows -Findings $findings `
        -Raw ([pscustomobject]@{
            LogName      = $logName
            LookbackDays = $days
            TotalEvents  = $totalCount
            CriticalEvents = $critCount
            GroupCount   = $groups.Count
            Truncated    = $truncated
            Groups       = @($groups | Select-Object -First $maxGroups)
        })
}

Register-PMCheck -Id 'SMB' -TitleKey 'smb.title' -Function 'Invoke-PMCheckSmbClientLog'
