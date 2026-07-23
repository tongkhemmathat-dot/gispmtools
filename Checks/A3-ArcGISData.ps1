# PMtools check - ArcGIS Server registered data store connections (enterprise
# geodatabases, file shares, cloud stores, big data file shares, NoSQL and
# raster stores), validated one by one.
# ASCII-only; text comes from i18n.json.
#
# Needs a configured connection and is disabled by default, like every other
# A*-ArcGIS* check - see A0-ArcGISSite.ps1.
#
# Uses two POST operations - /data/findItems and /data/validateDataItem -
# both POST because the Admin API requires it for these operations, not
# because either one writes. findItems only lists what is already
# registered; validateDataItem only asks the site to retest a connection it
# already has stored credentials for, using those stored credentials. No
# credential of any kind is supplied by this check, and none is stored.
#
# validateAllDataItems (the operation this check was originally planned
# around) turned out on a real site to answer only a bare {"status":
# "success"} for the whole site, with no per-item detail - useless for
# telling a healthy connection from a broken one among several. Calling
# validateDataItem once per item, discovered via findItems, gives that
# detail instead: each call answers "success" or "error" for exactly one
# connection.
#
# SECURITY: findItems returns each item's "info" block, which for an
# enterprise geodatabase includes a "connectionString" carrying the
# database server, database name, account name and an ArcGIS-encrypted
# password blob. That whole block is used ONLY in memory to build the
# validateDataItem request body for that same item, and is never placed in
# $rows, $findings, or -Raw - only Path, Type and the pass/fail outcome
# are. Grep PM-Data.json after changing this file if in doubt.

function Invoke-PMCheckArcGISData {

    try {
        $session = Get-PMArcGISSession

        $findResp = $null
        try {
            $findResp = Invoke-PMArcGISAdmin -Root $session.Root -Path 'data/findItems' -Token $session.Token `
                                              -TimeoutSec $session.TimeoutSec -Method Post `
                                              -Parameters @{ ancestorPath = '/'; types = 'egdb,folder,bigDataFileShare,nosql,cloudStore,rasterStore' }
        }
        catch {
            return New-PMResult -Id 'AGSDATA' -TitleKey 'agsdata.title' -Status 'WARN' `
                -SummaryKey 'agsdata.summary.listError' -SummaryValues @($_.Exception.Message) `
                -Findings @((New-PMFinding -Severity 'WARN' -TextKey 'agsdata.finding.listError' -Values @($_.Exception.Message)))
        }

        $items = @($findResp.items)
        if ($items.Count -eq 0) {
            return New-PMResult -Id 'AGSDATA' -TitleKey 'agsdata.title' -Status 'INFO' `
                -SummaryKey 'agsdata.summary.none'
        }

        $columns = @(
            (New-PMColumn -Key 'Path'   -TextKey 'agsdata.col.path' -Wide),
            (New-PMColumn -Key 'Type'   -TextKey 'agsdata.col.type'),
            (New-PMColumn -Key 'Status' -TextKey 'agsdata.col.status')
        )

        $rows      = @()
        $findings  = @()
        $failed    = 0

        foreach ($item in $items) {
            $path = [string]$item.path
            $type = [string]$item.type

            # Only Path, Type and Info travel into the validate request body -
            # Info never leaves this loop iteration.
            $itemJson = [pscustomobject]@{ path = $item.path; type = $item.type; info = $item.info } |
                        ConvertTo-Json -Depth 8 -Compress

            $status  = 'WARN'
            $message = ''
            try {
                $v = Invoke-PMArcGISAdmin -Root $session.Root -Path 'data/validateDataItem' -Token $session.Token `
                                           -TimeoutSec $session.TimeoutSec -Method Post -Parameters @{ item = $itemJson }
                if ([string]$v.status -eq 'success') {
                    $status = 'OK'
                }
                else {
                    if ($v.PSObject.Properties['message'])       { $message = [string]$v.message }
                    elseif ($v.PSObject.Properties['messages'])  { $message = (@($v.messages) -join '; ') }
                    elseif ($v.PSObject.Properties['error'])     { $message = [string]$v.error.message }
                }
            }
            catch { $message = $_.Exception.Message }

            if ($status -ne 'OK') {
                $failed++
                if ([string]::IsNullOrWhiteSpace($message)) {
                    $message = 'no error detail provided by the site'
                }
                $findings += New-PMFinding -Severity 'CRIT' -TextKey 'agsdata.finding.failed' -Values @($path, $message)
                $status = 'CRIT'
            }

            if ($status -eq 'OK') { $stateKey = 'agsdata.state.ok' } else { $stateKey = 'agsdata.state.failed' }
            $word = Get-PMWord -Key $stateKey
            $rows += @{
                Path       = $path
                Type       = $type
                Status     = $word.Th
                StatusEn   = $word.En
                _RowStatus = $status
            }
        }

        if ($failed -gt 0) {
            $status = 'CRIT'
            $sumKey = 'agsdata.summary.issue'
            $sumVal = @($items.Count, $failed)
        }
        else {
            $status = 'OK'
            $sumKey = 'agsdata.summary.ok'
            $sumVal = @($items.Count)
        }

        return New-PMResult -Id 'AGSDATA' -TitleKey 'agsdata.title' -Status $status `
            -SummaryKey $sumKey -SummaryValues $sumVal `
            -Columns $columns -Rows $rows -Findings $findings `
            -Raw ([pscustomobject]@{
                Total   = $items.Count
                Failed  = $failed
                Items   = @($rows | ForEach-Object { [pscustomobject]@{ Path = $_.Path; Type = $_.Type; Status = $_._RowStatus } })
            })
    }
    finally {
        Restore-PMArcGISCertificatePolicy
    }
}

Register-PMCheck -Id 'AGSDATA' -TitleKey 'agsdata.title' -Function 'Invoke-PMCheckArcGISData'
