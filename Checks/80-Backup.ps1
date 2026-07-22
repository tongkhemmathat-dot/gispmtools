# PMtools check - Data backup evidence.  ASCII-only; text comes from i18n.json.
#
# Two sources, because backup on Windows Server is done a dozen different ways:
#   1. the Microsoft-Windows-Backup log, which only exists once Windows Server
#      Backup is installed
#   2. scheduled tasks whose name matches Backup.TaskKeywords in settings.json,
#      which is how most third-party agents show up
# Finding nothing is reported as "no evidence found", not as "no backup exists":
# the backup may well be driven from an external system this check cannot see.

function Invoke-PMCheckBackup {

    $keywords = @(Get-PMSetting -Path 'Backup.TaskKeywords' -Default @('backup'))

    $columns = @(
        (New-PMColumn -Key 'Source' -TextKey 'backup.col.source'),
        (New-PMColumn -Key 'Name'   -TextKey 'backup.col.name' -Wide),
        (New-PMColumn -Key 'Last'   -TextKey 'backup.col.last'),
        (New-PMColumn -Key 'Result' -TextKey 'backup.col.result'),
        (New-PMColumn -Key 'Age'    -TextKey 'backup.col.age' -Align 'right')
    )

    $srcEvent = Get-PMWord -Key 'backup.src.eventlog'
    $srcTask  = Get-PMWord -Key 'backup.src.task'
    $wSuccess = Get-PMWord -Key 'backup.result.success'
    $wFailed  = Get-PMWord -Key 'backup.result.failed'

    $rows       = @()
    $raw        = @()
    $findings   = @()
    $newestOk   = $null
    $anyFailure = $null

    # --- source 1: Windows Server Backup event log ----------------------------
    $bkEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-Backup' } `
                               -MaxEvents 20 -ErrorAction SilentlyContinue)

    foreach ($e in $bkEvents) {
        # Event id 4 is the success record; every other id in this log is a
        # failure or an abort of some kind.
        $ok = ($e.Id -eq 4)
        if ($ok) { $w = $wSuccess; $status = 'OK' } else { $w = $wFailed; $status = 'CRIT' }

        $age = Get-PMAgeDays -Date $e.TimeCreated
        if ($ok -and ($null -eq $newestOk -or $e.TimeCreated -gt $newestOk)) { $newestOk = $e.TimeCreated }
        if (-not $ok -and $null -eq $anyFailure) { $anyFailure = $e.TimeCreated }

        $row = New-PMRow -Status $status -Values @{
            Source = $srcEvent.Th
            Name   = ("Event ID {0}" -f $e.Id)
            Last   = Format-PMStamp $e.TimeCreated
            Result = $w.Th
            Age    = $age
        }
        $row['SourceEn'] = $srcEvent.En
        $row['ResultEn'] = $w.En
        $rows += $row
        $raw  += [pscustomobject]@{ Source = 'EventLog'; EventId = $e.Id; Time = $e.TimeCreated; Success = $ok }
    }

    # --- source 2: scheduled tasks that look like backup jobs -----------------
    $tasks = @()
    try {
        $tasks = @(Get-ScheduledTask -ErrorAction Stop | Where-Object {
            $name = "$($_.TaskPath)$($_.TaskName)".ToLower()
            $hit = $false
            foreach ($k in $keywords) { if ($name -like "*$($k.ToLower())*") { $hit = $true; break } }
            # Exclude Microsoft's own housekeeping tasks, which match on the word
            # "backup" but say nothing about whether data is being protected.
            $hit -and ($_.TaskPath -notlike '\Microsoft\Windows\*')
        })
    }
    catch { $tasks = @() }

    foreach ($t in $tasks) {
        $info = $null
        try { $info = Get-ScheduledTaskInfo -TaskName $t.TaskName -TaskPath $t.TaskPath -ErrorAction Stop } catch { $info = $null }

        $lastRun = $null
        $ok      = $null
        if ($info) {
            if ($info.LastRunTime -and ([datetime]$info.LastRunTime).Year -gt 1900) { $lastRun = $info.LastRunTime }
            if ($null -ne $info.LastTaskResult) { $ok = ($info.LastTaskResult -eq 0) }
        }

        $age = Get-PMAgeDays -Date $lastRun
        if ($null -eq $ok)      { $status = 'INFO'; $resTh = '-';           $resEn = '' }
        elseif ($ok)            { $status = 'OK';   $resTh = $wSuccess.Th;  $resEn = $wSuccess.En }
        else                    { $status = 'CRIT'; $resTh = $wFailed.Th;   $resEn = $wFailed.En }

        if ($ok -and $lastRun -and ($null -eq $newestOk -or $lastRun -gt $newestOk)) { $newestOk = $lastRun }
        if (($ok -eq $false) -and $null -eq $anyFailure) { $anyFailure = $t.TaskName }

        $row = New-PMRow -Status $status -Values @{
            Source = $srcTask.Th
            Name   = "$($t.TaskPath)$($t.TaskName)"
            Last   = $(if ($lastRun) { Format-PMStamp $lastRun } else { '-' })
            Result = $resTh
            Age    = $(if ($null -ne $age) { $age } else { '-' })
        }
        $row['SourceEn'] = $srcTask.En
        $row['ResultEn'] = $resEn
        $rows += $row
        $raw  += [pscustomobject]@{ Source = 'ScheduledTask'; Name = "$($t.TaskPath)$($t.TaskName)"; LastRun = $lastRun; Success = $ok }
    }

    # --- verdict --------------------------------------------------------------
    if ($rows.Count -eq 0) {
        $findings += New-PMFinding -Severity 'WARN' -TextKey 'backup.finding.none'
        return New-PMResult -Id 'BACKUP' -TitleKey 'backup.title' -Status 'WARN' `
            -SummaryKey 'backup.summary.none' `
            -Columns $columns -Rows @() -Findings $findings `
            -Raw ([pscustomobject]@{ EvidenceFound = $false })
    }

    $status = 'OK'

    if ($null -ne $anyFailure) {
        $status = 'CRIT'
        $findings += New-PMFinding -Severity 'CRIT' -TextKey 'backup.finding.failed' -Values @((Format-PMStamp $anyFailure))
    }

    if ($null -eq $newestOk) {
        $status = Get-PMWorstStatus @($status, 'WARN')
        $sumKey = 'backup.summary.issue'; $sumVal = @()
    }
    else {
        $ageDays  = [math]::Floor(((Get-Date) - [datetime]$newestOk).TotalDays)
        $ageState = Test-PMThreshold -Name 'BackupAgeDays' -Value $ageDays
        if ($ageState -eq 'WARN' -or $ageState -eq 'CRIT') {
            $status = Get-PMWorstStatus @($status, $ageState)
            $findings += New-PMFinding -Severity $ageState -TextKey 'backup.finding.old' `
                -Values @($ageDays, (Get-PMThresholdValue -Name 'BackupAgeDays' -Level 'Warn'))
            $sumKey = 'backup.summary.issue'; $sumVal = @()
        }
        else {
            $sumKey = 'backup.summary.ok'; $sumVal = @((Format-PMStamp $newestOk), $ageDays)
        }
    }

    return New-PMResult -Id 'BACKUP' -TitleKey 'backup.title' -Status $status `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns $columns -Rows $rows -Findings $findings `
        -Raw ([pscustomobject]@{ EvidenceFound = $true; NewestSuccess = $newestOk; Entries = $raw })
}

Register-PMCheck -Id 'BACKUP' -TitleKey 'backup.title' -Function 'Invoke-PMCheckBackup'
