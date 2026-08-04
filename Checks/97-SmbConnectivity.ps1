# PMtools check - SMB/NAS connectivity: can this server reach a configured
# NAS on port 445, and if a share path was given, can it actually be read?
# ASCII-only; text comes from i18n.json.
#
# DISABLED BY DEFAULT in Config\settings.json (Checks.Disabled), for the same
# reason as CONN (Checks\95-Connectivity.ps1): this is the one other check
# that reaches off the box. No targets are known in advance either - there is
# no way to guess which NAS a given server depends on - so Config\settings.json
# SmbConnectivity.Targets starts empty and this check reports INFO "nothing
# configured" until an admin adds entries.
#
# READ-ONLY: the share test is Test-Path on the UNC path, which is a metadata
# read (equivalent to an SMB QUERY_DIRECTORY/QUERY_INFO), not a file open for
# write. Nothing is created, changed, or left behind on the NAS.
#
# The share is only probed after the port test succeeds - Test-Path against
# an unreachable host can hang for the OS's own SMB connect timeout (which is
# much longer, and not configurable per-call), so a dead NAS would otherwise
# turn one bad target into a multi-second stall per check run.
#
# Test-PMSmbPort lives in Lib\Core.ps1, not here - the "Test NAS / SMB
# connectivity" menu (Show-PMMenu.ps1) needs the same fast-timeout port test
# for its ad-hoc/interval mode, and Core.ps1 is always loaded before either.

function Invoke-PMCheckSmbConnectivity {

    $timeoutMs = [int](Get-PMSetting -Path 'SmbConnectivity.TimeoutMs' -Default 1500)
    $targets   = @(Get-PMSetting -Path 'SmbConnectivity.Targets' -Default @())

    $columns = @(
        (New-PMColumn -Key 'Name'    -TextKey 'smbconn.col.name'),
        (New-PMColumn -Key 'Server'  -TextKey 'smbconn.col.server'),
        (New-PMColumn -Key 'Port'    -TextKey 'smbconn.col.port'),
        (New-PMColumn -Key 'Share'   -TextKey 'smbconn.col.share'),
        (New-PMColumn -Key 'Latency' -TextKey 'smbconn.col.latency' -Align 'right')
    )

    if ($targets.Count -eq 0) {
        return New-PMResult -Id 'SMBCONN' -TitleKey 'smbconn.title' -Status 'INFO' `
            -SummaryKey 'smbconn.summary.none' `
            -Columns $columns -Rows @() -Findings @() `
            -Raw ([pscustomobject]@{ TargetCount = 0 })
    }

    $reachWord    = Get-PMWord -Key 'smbconn.result.reachable'
    $unreachWord  = Get-PMWord -Key 'smbconn.result.unreachable'
    $accessWord   = Get-PMWord -Key 'smbconn.result.accessible'
    $noaccessWord = Get-PMWord -Key 'smbconn.result.notaccessible'

    $rows     = @()
    $findings = @()
    $raw      = @()
    $failed   = 0

    foreach ($t in $targets) {
        $server = [string]$t.Server
        $name   = [string]$t.Name
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $server }
        $port      = 445
        if ($t.Port) { $port = [int]$t.Port }
        $sharePath = [string]$t.SharePath

        if ([string]::IsNullOrWhiteSpace($server)) { continue }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $portOk = Test-PMSmbPort -ComputerName $server -Port $port -TimeoutMs $timeoutMs
        $sw.Stop()
        $latency = if ($portOk) { [int]$sw.ElapsedMilliseconds } else { '-' }

        $shareOk = $null   # $null = not tested (no SharePath given, or port already failed)
        if ($portOk -and -not [string]::IsNullOrWhiteSpace($sharePath)) {
            try { $shareOk = [bool](Test-Path -LiteralPath $sharePath -ErrorAction Stop) }
            catch { $shareOk = $false }
        }

        if (-not $portOk) {
            $status   = 'WARN'
            $portTh   = $unreachWord.Th;  $portEn = $unreachWord.En
            $shareTh  = '-'; $shareEn = '-'
            $failed++
            $findings += New-PMFinding -Severity 'WARN' -TextKey 'smbconn.finding.port' -Values @($name, $server)
        }
        else {
            $portTh = $reachWord.Th; $portEn = $reachWord.En
            if ($null -eq $shareOk) {
                $shareTh = '-'; $shareEn = '-'
                $status  = 'OK'
            }
            elseif ($shareOk) {
                $shareTh = $accessWord.Th; $shareEn = $accessWord.En
                $status  = 'OK'
            }
            else {
                $shareTh = $noaccessWord.Th; $shareEn = $noaccessWord.En
                $status  = 'WARN'
                $failed++
                $findings += New-PMFinding -Severity 'WARN' -TextKey 'smbconn.finding.share' -Values @($server, $sharePath)
            }
        }

        $row = New-PMRow -Status $status -Values @{
            Name    = $name
            Server  = $server
            Port    = $portTh
            Share   = $shareTh
            Latency = $latency
        }
        $row['PortEn']  = $portEn
        $row['ShareEn'] = $shareEn
        $rows += $row

        $raw += [pscustomobject]@{
            Name       = $name
            Server     = $server
            Port       = $port
            PortOk     = $portOk
            SharePath  = $sharePath
            ShareOk    = $shareOk
            LatencyMs  = $latency
        }
    }

    if ($rows.Count -eq 0) {
        return New-PMResult -Id 'SMBCONN' -TitleKey 'smbconn.title' -Status 'INFO' `
            -SummaryKey 'smbconn.summary.none' `
            -Columns $columns -Rows @() -Findings @() `
            -Raw ([pscustomobject]@{ TargetCount = 0 })
    }

    if ($failed -gt 0) { $status = 'WARN'; $sumKey = 'smbconn.summary.issue'; $sumVal = @($rows.Count, $failed) }
    else               { $status = 'OK';   $sumKey = 'smbconn.summary.ok';    $sumVal = @($rows.Count) }

    return New-PMResult -Id 'SMBCONN' -TitleKey 'smbconn.title' -Status $status `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns $columns -Rows $rows -Findings $findings -Raw $raw
}

Register-PMCheck -Id 'SMBCONN' -TitleKey 'smbconn.title' -Function 'Invoke-PMCheckSmbConnectivity'
