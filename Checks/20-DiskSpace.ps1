# PMtools check - Fixed disk free space.  ASCII-only; text comes from i18n.json.

function Invoke-PMCheckDiskSpace {

    $disks = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)

    $columns = @(
        (New-PMColumn -Key 'Drive'      -TextKey 'disk.col.drive'),
        (New-PMColumn -Key 'Label'      -TextKey 'disk.col.label'),
        (New-PMColumn -Key 'FileSystem' -TextKey 'disk.col.fs'),
        (New-PMColumn -Key 'TotalGB'    -TextKey 'disk.col.total'   -Align 'right'),
        (New-PMColumn -Key 'UsedGB'     -TextKey 'disk.col.used'    -Align 'right'),
        (New-PMColumn -Key 'FreeGB'     -TextKey 'disk.col.free'    -Align 'right'),
        (New-PMColumn -Key 'FreePct'    -TextKey 'disk.col.freepct' -Align 'right')
    )

    $rows     = @()
    $findings = @()
    $raw      = @()
    $issues   = 0

    foreach ($d in ($disks | Sort-Object DeviceID)) {

        if (-not $d.Size -or [double]$d.Size -le 0) { continue }   # unformatted / unavailable volume

        $totalGB = ConvertTo-PMGB $d.Size
        $freeGB  = ConvertTo-PMGB $d.FreeSpace
        $usedGB  = [math]::Round($totalGB - $freeGB, 2)
        $freePct = [math]::Round((([double]$d.FreeSpace / [double]$d.Size) * 100), 1)

        # A drive is judged on both measures: a 5% shortfall on a 4 TB volume is
        # still hundreds of GB, and 15 GB free on a small volume is still tight.
        $status = Get-PMWorstStatus @(
            (Test-PMThreshold -Name 'DiskFreePercent' -Value $freePct),
            (Test-PMThreshold -Name 'DiskFreeGB'      -Value $freeGB)
        )
        if ($status -eq 'WARN' -or $status -eq 'CRIT') { $issues++ }

        $rows += New-PMRow -Status $status -Values @{
            Drive      = $d.DeviceID
            Label      = $d.VolumeName
            FileSystem = $d.FileSystem
            TotalGB    = $totalGB
            UsedGB     = $usedGB
            FreeGB     = $freeGB
            FreePct    = $freePct
        }

        if ($status -eq 'CRIT') {
            $findings += New-PMFinding -Severity 'CRIT' -TextKey 'disk.finding.crit' -Values @($d.DeviceID, $freePct, $freeGB)
        }
        elseif ($status -eq 'WARN') {
            $findings += New-PMFinding -Severity 'WARN' -TextKey 'disk.finding.warn' -Values @($d.DeviceID, $freePct, $freeGB)
        }

        $raw += [pscustomobject]@{
            Drive = $d.DeviceID; Label = $d.VolumeName; FileSystem = $d.FileSystem
            TotalGB = $totalGB; UsedGB = $usedGB; FreeGB = $freeGB; FreePercent = $freePct; Status = $status
        }
    }

    $overall = Get-PMWorstStatus (@($rows | ForEach-Object { $_._RowStatus }) + @('OK'))

    if ($issues -gt 0) { $sumKey = 'disk.summary.issue'; $sumVal = @($rows.Count, $issues) }
    else               { $sumKey = 'disk.summary.ok';    $sumVal = @($rows.Count) }

    return New-PMResult -Id 'DISK' -TitleKey 'disk.title' -Status $overall `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns $columns -Rows $rows -Findings $findings -Raw $raw
}

Register-PMCheck -Id 'DISK' -TitleKey 'disk.title' -Function 'Invoke-PMCheckDiskSpace'
