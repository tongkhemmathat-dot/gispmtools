# PMtools check - Pending restart indicators.  ASCII-only; text comes from i18n.json.
#
# Windows records a pending restart in several unrelated places depending on
# what asked for it, so all four are read and reported individually. A server
# that has been waiting to restart for weeks is patched on paper but not in fact.

function Invoke-PMCheckPendingReboot {

    $indicators = @(
        @{ Key = 'reboot.item.cbs'
           Test = { Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' } }

        @{ Key = 'reboot.item.wu'
           Test = { Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' } }

        @{ Key = 'reboot.item.rename'
           Test = {
               $v = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' `
                                     -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
               return ($null -ne $v -and $null -ne $v.PendingFileRenameOperations -and @($v.PendingFileRenameOperations).Count -gt 0)
           } }

        @{ Key = 'reboot.item.netlogon'
           Test = {
               (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services\Pending') -or
               (Test-Path -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\JoinDomain') -or
               (Test-Path -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\AvoidSpnSet')
           } }
    )

    $pendingWord = Get-PMWord -Key 'reboot.value.pending'
    $clearWord   = Get-PMWord -Key 'reboot.value.clear'

    $rows        = @()
    $raw         = @()
    $pendingList = @()

    foreach ($ind in $indicators) {
        $isPending = $false
        try { $isPending = [bool](& $ind.Test) } catch { $isPending = $false }

        if ($isPending) { $status = 'WARN'; $w = $pendingWord } else { $status = 'OK'; $w = $clearWord }

        $label = Get-PMText -Key $ind.Key
        if ($isPending) { $pendingList += $label.Th }

        $rows += New-PMItemRow -TextKey $ind.Key -Value $w.Th -ValueEn $w.En -Status $status
        $raw  += [pscustomobject]@{ Indicator = $label.En; Pending = $isPending }
    }

    $findings = @()
    if ($pendingList.Count -gt 0) {
        $status = 'WARN'
        $sumKey = 'reboot.summary.issue'; $sumVal = @($pendingList.Count)
        $findings += New-PMFinding -Severity 'WARN' -TextKey 'reboot.finding.pending' -Values @(($pendingList -join ', '))
    }
    else {
        $status = 'OK'
        $sumKey = 'reboot.summary.ok'; $sumVal = @()
    }

    return New-PMResult -Id 'REBOOT' -TitleKey 'reboot.title' -Status $status `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns (New-PMItemColumns) -Rows $rows -Findings $findings `
        -Raw ([pscustomobject]@{ PendingCount = $pendingList.Count; Indicators = $raw })
}

Register-PMCheck -Id 'REBOOT' -TitleKey 'reboot.title' -Function 'Invoke-PMCheckPendingReboot'
