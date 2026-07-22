# PMtools check - Time synchronisation.  ASCII-only; text comes from i18n.json.
#
# w32tm prints localised labels and returns a non-zero exit code in perfectly
# normal situations, so neither the label text nor $LASTEXITCODE can be trusted.
# What is stable across languages is the line order and the numeric values:
# line 1 is always the leap indicator and line 2 the stratum. Where even that
# cannot be read the check reports INFO with the raw output rather than
# inventing a verdict.

function Invoke-PMCheckTimeSync {

    $statusText = @()
    $sourceText = ''
    try { $statusText = @(& w32tm /query /status 2>&1 | ForEach-Object { [string]$_ }) } catch { $statusText = @() }
    try { $sourceText = (& w32tm /query /source 2>&1 | Select-Object -First 1) -as [string] } catch { $sourceText = '' }

    $leap    = $null
    $stratum = $null
    if ($statusText.Count -ge 2) {
        if ($statusText[0] -match ':\s*(\d+)') { $leap    = [int]$Matches[1] }
        if ($statusText[1] -match ':\s*(\d+)') { $stratum = [int]$Matches[1] }
    }

    # Leap indicator 3 means "clock not synchronised"; stratum 0 means the
    # source is unspecified. Either one on its own means we are not in sync.
    $synced = $null
    if ($null -ne $leap -and $null -ne $stratum) {
        $synced = (($leap -ne 3) -and ($stratum -gt 0))
    }

    $svc      = Get-Service -Name 'W32Time' -ErrorAction SilentlyContinue
    $svcState = $null
    if ($svc) { $svcState = ($svc.Status -eq 'Running') }

    $lastSync = ''
    foreach ($line in $statusText) {
        if ($line -match '(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s+\d{1,2}:\d{2}:\d{2})') { $lastSync = $Matches[1]; break }
    }
    $offset = ''
    foreach ($line in $statusText) {
        if ($line -match '([-+]?\d+\.\d+s)') { $offset = $Matches[1]; break }
    }

    $enabled  = Get-PMWord -Key 'common.enabled'
    $disabled = Get-PMWord -Key 'common.disabled'
    $unknown  = Get-PMWord -Key 'common.unknown'

    $findings = @()

    if ($null -eq $synced) {
        $syncStatus = 'INFO'
        $syncTh = $unknown.Th; $syncEn = $unknown.En
    }
    elseif ($synced) {
        $syncStatus = 'OK'
        $syncTh = $enabled.Th; $syncEn = $enabled.En
    }
    else {
        $syncStatus = 'CRIT'
        $syncTh = $disabled.Th; $syncEn = $disabled.En
        $findings += New-PMFinding -Severity 'CRIT' -TextKey 'time.finding.nosync'
    }

    $rows = @(
        (New-PMItemRow -TextKey 'time.item.localtime' -Value ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')))
        (New-PMItemRow -TextKey 'time.item.source'    -Value $sourceText)
        (New-PMItemRow -TextKey 'time.item.leap'      -Value $syncTh -ValueEn $syncEn -Status $syncStatus)
    )
    if ($null -ne $stratum) { $rows += New-PMItemRow -TextKey 'time.item.stratum' -Value $stratum }
    if ($lastSync)          { $rows += New-PMItemRow -TextKey 'time.item.lastsync' -Value $lastSync }
    if ($offset)            { $rows += New-PMItemRow -TextKey 'time.item.offset'   -Value $offset }

    $svcStatus = 'INFO'
    if ($null -ne $svcState) {
        if ($svcState) { $svcStatus = 'OK'; $v = $enabled } else { $svcStatus = 'WARN'; $v = $disabled }
        $rows += New-PMItemRow -TextKey 'time.item.service' -Value $v.Th -ValueEn $v.En -Status $svcStatus
        if (-not $svcState) { $findings += New-PMFinding -Severity 'WARN' -TextKey 'time.finding.svcstopped' }
    }

    $overall = Get-PMWorstStatus @($syncStatus, $svcStatus)

    if ($syncStatus -eq 'OK') { $sumKey = 'time.summary.ok'; $sumVal = @($sourceText) }
    else                      { $sumKey = 'time.summary.issue'; $sumVal = @() }

    return New-PMResult -Id 'TIME' -TitleKey 'time.title' -Status $overall `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns (New-PMItemColumns) -Rows $rows -Findings $findings `
        -Raw ([pscustomobject]@{
            Source = $sourceText; LeapIndicator = $leap; Stratum = $stratum
            Synchronized = $synced; ServiceRunning = $svcState; RawStatus = $statusText
        })
}

Register-PMCheck -Id 'TIME' -TitleKey 'time.title' -Function 'Invoke-PMCheckTimeSync'
