# PMtools check - Pending Windows updates.  ASCII-only; text comes from i18n.json.
#
# DISABLED BY DEFAULT in Config\settings.json (Checks.Disabled).
#
# The search is still read-only - it asks what is pending and installs nothing -
# but it contacts WSUS or Microsoft Update and can run for many minutes or hang
# outright on an isolated network, which is not acceptable in a routine PM pass.
# Run it deliberately when it is wanted:  .\Start-PMCheck.ps1 -Only WU

function Invoke-PMCheckWindowsUpdate {

    $columns = @(
        (New-PMColumn -Key 'Title'    -TextKey 'wu.col.title' -Wide),
        (New-PMColumn -Key 'Kb'       -TextKey 'wu.col.kb'),
        (New-PMColumn -Key 'Category' -TextKey 'wu.col.category'),
        (New-PMColumn -Key 'Severity' -TextKey 'wu.col.severity'),
        (New-PMColumn -Key 'Size'     -TextKey 'wu.col.size' -Align 'right')
    )

    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $query    = "IsInstalled=0 and IsHidden=0"

    $result = $null
    $online = $true
    try {
        $result = $searcher.Search($query)
    }
    catch {
        # No route to the update source. Fall back to the locally cached result
        # so the report still says something useful, and mark it as offline.
        $online = $false
        try {
            $searcher.Online = $false
            $result = $searcher.Search($query)
        }
        catch {
            return New-PMResult -Id 'WU' -TitleKey 'wu.title' -Status 'WARN' `
                -SummaryKey 'wu.unavailable' `
                -Findings @(New-PMFinding -Severity 'WARN' -TextKey 'wu.unavailable') `
                -Raw ([pscustomobject]@{ Error = $_.Exception.Message })
        }
    }

    $rows        = @()
    $raw         = @()
    $securityCnt = 0

    foreach ($u in $result.Updates) {

        $categories = @()
        foreach ($c in $u.Categories) { $categories += $c.Name }

        $kbs = @()
        foreach ($k in $u.KBArticleIDs) { $kbs += "KB$k" }

        $severity = $u.MsrcSeverity
        $isSecurity = ($categories -contains 'Security Updates') -or (-not [string]::IsNullOrWhiteSpace($severity))
        if ($isSecurity) { $securityCnt++; $status = 'CRIT' } else { $status = 'WARN' }

        $sizeMB = [math]::Round(([double]$u.MaxDownloadSize / 1MB), 1)

        $rows += New-PMRow -Status $status -Values @{
            Title    = Get-PMShortText -Text $u.Title -MaxLength 150
            Kb       = ($kbs -join ', ')
            Category = ($categories -join ', ')
            Severity = $severity
            Size     = $sizeMB
        }
        $raw += [pscustomobject]@{
            Title = $u.Title; KB = $kbs; Categories = $categories
            MsrcSeverity = $severity; SizeMB = $sizeMB; IsSecurity = $isSecurity
        }
    }

    $total    = $rows.Count
    $findings = @()

    if ($total -eq 0) {
        $status = 'OK'; $sumKey = 'wu.summary.ok'; $sumVal = @()
    }
    else {
        $sumKey = 'wu.summary.issue'; $sumVal = @($total, $securityCnt)
        if ($securityCnt -gt 0) {
            $status = 'CRIT'
            $findings += New-PMFinding -Severity 'CRIT' -TextKey 'wu.finding.security' -Values @($securityCnt)
        }
        else {
            $status = 'WARN'
            $findings += New-PMFinding -Severity 'WARN' -TextKey 'wu.finding.pending' -Values @($total)
        }
    }
    if (-not $online) {
        $findings += New-PMFinding -Severity 'INFO' -TextKey 'wu.unavailable'
    }

    return New-PMResult -Id 'WU' -TitleKey 'wu.title' -Status $status `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns $columns -Rows $rows -Findings $findings `
        -Raw ([pscustomobject]@{ OnlineSearch = $online; PendingCount = $total; SecurityCount = $securityCnt; Updates = $raw })
}

Register-PMCheck -Id 'WU' -TitleKey 'wu.title' -Function 'Invoke-PMCheckWindowsUpdate'
