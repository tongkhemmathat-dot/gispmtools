# PMtools check - Recently installed updates.  ASCII-only; text comes from i18n.json.
# Evidence that patching is actually happening, independent of what is pending.

function Invoke-PMCheckHotfix {

    $take = [int](Get-PMSetting -Path 'Lookback.HotfixCount' -Default 10)

    $hotfixes = @(Get-HotFix -ErrorAction Stop |
                  Sort-Object -Property @{ Expression = { $_.InstalledOn }; Descending = $true })

    $columns = @(
        (New-PMColumn -Key 'Kb'        -TextKey 'hotfix.col.kb'),
        (New-PMColumn -Key 'Desc'      -TextKey 'hotfix.col.desc'),
        (New-PMColumn -Key 'Installed' -TextKey 'hotfix.col.installed'),
        (New-PMColumn -Key 'By'        -TextKey 'hotfix.col.by' -Wide)
    )

    $rows = @()
    $raw  = @()
    foreach ($h in ($hotfixes | Select-Object -First $take)) {
        $rows += New-PMRow -Values @{
            Kb        = $h.HotFixID
            Desc      = $h.Description
            Installed = $(if ($h.InstalledOn) { ([datetime]$h.InstalledOn).ToString('yyyy-MM-dd') } else { '-' })
            By        = $h.InstalledBy
        }
        $raw += [pscustomobject]@{ HotFixID = $h.HotFixID; Description = $h.Description; InstalledOn = $h.InstalledOn; InstalledBy = $h.InstalledBy }
    }

    $findings = @()

    # InstalledOn is null for some entries (notably OEM-injected packages), so
    # take the newest date that actually exists rather than the newest row.
    $latest = ($hotfixes | Where-Object { $_.InstalledOn } |
               Measure-Object -Property InstalledOn -Maximum).Maximum

    if ($null -eq $latest) {
        $status = 'WARN'
        $sumKey = 'hotfix.summary.none'; $sumVal = @()
        $findings += New-PMFinding -Severity 'WARN' -TextKey 'hotfix.finding.none'
    }
    else {
        $ageDays = [math]::Floor(((Get-Date) - [datetime]$latest).TotalDays)
        $status  = Test-PMThreshold -Name 'LastPatchDays' -Value $ageDays
        $sumKey  = 'hotfix.summary'; $sumVal = @(([datetime]$latest).ToString('yyyy-MM-dd'), $ageDays)

        if ($status -eq 'WARN' -or $status -eq 'CRIT') {
            $findings += New-PMFinding -Severity $status -TextKey 'hotfix.finding.stale' `
                -Values @($ageDays, (Get-PMThresholdValue -Name 'LastPatchDays' -Level 'Warn'))
        }
    }

    return New-PMResult -Id 'HOTFIX' -TitleKey 'hotfix.title' -Status $status `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns $columns -Rows $rows -Findings $findings `
        -Raw ([pscustomobject]@{ TotalInstalled = $hotfixes.Count; LatestInstall = $latest; Recent = $raw })
}

Register-PMCheck -Id 'HOTFIX' -TitleKey 'hotfix.title' -Function 'Invoke-PMCheckHotfix'
