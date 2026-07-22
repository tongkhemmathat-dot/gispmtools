# PMtools check - Antivirus / antimalware state.  ASCII-only; text comes from i18n.json.
#
# Three sources are tried in order, because none of them works everywhere:
#   1. Get-MpComputerStatus  - richest data, but only where Defender is present
#   2. root\SecurityCenter2  - covers third-party products, but this WMI
#                              namespace does NOT exist on Windows Server
#   3. the WinDefend service - last resort, state only
# If all three come up empty the check says so plainly rather than implying the
# server is unprotected or, worse, implying that it is protected.

function Invoke-PMCheckAntivirus {

    $rows     = @()
    $findings = @()
    $source   = $null
    $product  = $null
    $rtp      = $null
    $sigDate  = $null
    $sigVer   = $null
    $lastScan = $null
    $svcOk    = $null

    $enabled  = Get-PMWord -Key 'common.enabled'
    $disabled = Get-PMWord -Key 'common.disabled'
    $unknown  = Get-PMWord -Key 'common.unknown'

    # --- source 1: Defender ---------------------------------------------------
    $mp = $null
    try { $mp = Get-MpComputerStatus -ErrorAction Stop } catch { $mp = $null }

    if ($null -ne $mp) {
        $source   = 'Get-MpComputerStatus'
        $product  = 'Microsoft Defender Antivirus'
        $rtp      = [bool]$mp.RealTimeProtectionEnabled
        $svcOk    = [bool]$mp.AMServiceEnabled
        $sigDate  = $mp.AntivirusSignatureLastUpdated
        $sigVer   = $mp.AntivirusSignatureVersion
        $lastScan = $mp.FullScanEndTime
    }
    else {
        # --- source 2: Security Center (client OS only) -----------------------
        $sc = @()
        try { $sc = @(Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction Stop) } catch { $sc = @() }

        if ($sc.Count -gt 0) {
            $av = $sc[0]
            $source  = 'root\SecurityCenter2'
            $product = $av.displayName
            # productState is a packed bitfield; byte 2 carries real-time status
            # (0x10/0x11 = on) and byte 3 the definition state (0x00 = current).
            $hex = '{0:x6}' -f [int]$av.productState
            $rtp = ($hex.Substring(2, 2) -eq '10' -or $hex.Substring(2, 2) -eq '11')
            if ($av.timestamp) { try { $sigDate = [datetime]::Parse($av.timestamp) } catch { $sigDate = $null } }
        }
        else {
            # --- source 3: the Defender service alone -------------------------
            $svc = Get-Service -Name 'WinDefend' -ErrorAction SilentlyContinue
            if ($svc) {
                $source  = 'Service: WinDefend'
                $product = 'Microsoft Defender Antivirus'
                $svcOk   = ($svc.Status -eq 'Running')
            }
        }
    }

    if ($null -eq $source) {
        $findings += New-PMFinding -Severity 'CRIT' -TextKey 'av.finding.notfound'
        return New-PMResult -Id 'AV' -TitleKey 'av.title' -Status 'CRIT' `
            -SummaryKey 'av.summary.notfound' `
            -Columns (New-PMItemColumns) -Rows @() -Findings $findings `
            -Raw ([pscustomobject]@{ Detected = $false })
    }

    # --- build the report rows -----------------------------------------------
    $rows += New-PMItemRow -TextKey 'av.item.product' -Value $product
    $rows += New-PMItemRow -TextKey 'av.item.source'  -Value $source

    $rtpStatus = 'INFO'
    if ($null -ne $rtp) {
        if ($rtp) { $rtpStatus = 'OK';   $v = $enabled }
        else      { $rtpStatus = 'CRIT'; $v = $disabled }
        $rows += New-PMItemRow -TextKey 'av.item.rtp' -Value $v.Th -ValueEn $v.En -Status $rtpStatus
        if (-not $rtp) { $findings += New-PMFinding -Severity 'CRIT' -TextKey 'av.finding.rtpoff' }
    }

    if ($null -ne $svcOk) {
        if ($svcOk) { $v = $enabled } else { $v = $disabled }
        $rows += New-PMItemRow -TextKey 'av.item.service' -Value $v.Th -ValueEn $v.En `
            -Status $(if ($svcOk) { 'OK' } else { 'CRIT' })
        if (-not $svcOk) { $rtpStatus = Get-PMWorstStatus @($rtpStatus, 'CRIT') }
    }

    if ($sigVer) { $rows += New-PMItemRow -TextKey 'av.item.sigversion' -Value $sigVer }

    $sigStatus = 'INFO'
    $sigDisplay = $unknown.Th
    if ($null -ne $sigDate) {
        $ageDays    = Get-PMAgeDays -Date $sigDate
        $sigStatus  = Test-PMThreshold -Name 'AvSignatureDays' -Value $ageDays
        $sigDisplay = Format-PMStamp $sigDate
        $rows += New-PMItemRow -TextKey 'av.item.sigdate' -Value $sigDisplay
        $rows += New-PMItemRow -TextKey 'av.item.sigage'  -Value ((Get-PMText -Key 'common.days' -Values @($ageDays)).Th) -Status $sigStatus

        if ($sigStatus -eq 'WARN' -or $sigStatus -eq 'CRIT') {
            $findings += New-PMFinding -Severity $sigStatus -TextKey 'av.finding.sigold' `
                -Values @($ageDays, (Get-PMThresholdValue -Name 'AvSignatureDays' -Level 'Warn'))
        }
    }

    if ($lastScan) { $rows += New-PMItemRow -TextKey 'av.item.lastscan' -Value (Format-PMStamp $lastScan) }

    $overall = Get-PMWorstStatus @($rtpStatus, $sigStatus, 'OK')

    if ($overall -eq 'OK') { $sumKey = 'av.summary.ok'; $sumVal = @($product, $sigDisplay) }
    else                   { $sumKey = 'av.summary.issue'; $sumVal = @($product) }

    return New-PMResult -Id 'AV' -TitleKey 'av.title' -Status $overall `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns (New-PMItemColumns) -Rows $rows -Findings $findings `
        -Raw ([pscustomobject]@{
            Detected = $true; Source = $source; Product = $product
            RealTimeProtection = $rtp; ServiceEnabled = $svcOk
            SignatureVersion = $sigVer; SignatureDate = $sigDate; LastFullScan = $lastScan
        })
}

Register-PMCheck -Id 'AV' -TitleKey 'av.title' -Function 'Invoke-PMCheckAntivirus'
