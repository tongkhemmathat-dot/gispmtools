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
# SECURITY: findItems returns each item's "info" block, which carries
# secrets - a "connectionString" with an ArcGIS-encrypted password blob for
# an enterprise geodatabase, "accessKey"/"secretKey" for an objectStore,
# a "users[].password" list for a nosql store. That whole block is used
# ONLY in memory to build a request for that same item (the
# validateDataItem body below, or - for nosql/objectStore - to read the
# machine name(s) it already lists at info.machines[].name so their own
# health-check operation can be called) and is never placed in $rows,
# $findings, or -Raw - only Path, Type and the pass/fail outcome are.
# Grep PM-Data.json after changing this file if in doubt.
#
# validateDataItem cannot be called for every type findItems can return.
# Confirmed against a real ArcGIS Data Store-backed hosting server (2026-07-
# 24): "nosql" and "objectStore" items - both internal to ArcGIS Data Store
# itself (replication log, tile cache, object store; named "AGSDataStore_*"
# by the server, not by an administrator) - always answer validateDataItem
# with a raw Java NullPointerException from the site's own admin/dataspace
# code (DataItem.getProvider() returning null), never a real success/
# failure verdict. That is a site-side gap in validateDataItem's support
# for its own managed item types, not a connection problem.
#
# ArcGIS Data Store items have their OWN, different health-check operation
# instead: POST data/items/<path>/machines/<machineName>/validate, which
# answers {"status":"success","datastore.overallhealth":"Healthy",
# "machines":[{"machine.overallhealth":"Healthy",...}],...} - confirmed
# against the same real site (2026-07-24), including for the exact
# objectStore item validateDataItem NPEs on. machineName is not guessed: it
# comes straight from that item's own info.machines[].name, already in
# memory from findItems.
#
# "nosql" is skipped outright rather than run through that same operation.
# Confirmed on the same real site that "nosql" covers two very different
# sub-types with two very different results: an Apache Ignite-backed
# in-memory cache item answered a real, correct CRIT (the cache feature
# genuinely was not installed on that machine), while a RabbitMQ-backed
# queue item answered the site's standard "no such resource" text because
# the operation is not exposed for that sub-type at all. Telling those
# apart reliably from outside would mean matching on implementation details
# (Ignite vs RabbitMQ, "cacheStore" vs "queueStore") that are not
# guaranteed stable across ArcGIS Data Store versions, so by request "nosql"
# items are reported as skipped instead of risking either a false CRIT on
# a queue-type item elsewhere, or silently trusting a health result this
# check cannot tell apart from "operation not found" reliably enough.
# "objectStore" did not show this ambiguity (Ozone-backed, single behavior
# observed) and keeps validating for real. "egdb" (the actual database
# hosted feature layers query) uses validateDataItem as normal and is the
# type that matters for this check regardless of whether it also carries
# an "AGSDataStore_" name.

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

        $rows            = @()
        $findings        = @()
        $failed          = 0
        $skipped         = 0
        $skipTypes       = @('nosql')
        $dataStoreTypes  = @('objectStore')

        foreach ($item in $items) {
            $path = [string]$item.path
            $type = [string]$item.type

            if ($skipTypes -contains $type) {
                $skipped++
                $findings += New-PMFinding -Severity 'INFO' -TextKey 'agsdata.finding.skipped' -Values @($path, $type)
                $word = Get-PMWord -Key 'agsdata.state.skipped'
                $rows += @{
                    Path       = $path
                    Type       = $type
                    Status     = $word.Th
                    StatusEn   = $word.En
                    _RowStatus = 'INFO'
                }
                continue
            }

            if ($dataStoreTypes -contains $type) {
                $machineNames = @($item.info.machines | ForEach-Object { [string]$_.name } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

                if ($machineNames.Count -eq 0) {
                    $skipped++
                    $findings += New-PMFinding -Severity 'INFO' -TextKey 'agsdata.finding.skipped' -Values @($path, $type)
                    $word = Get-PMWord -Key 'agsdata.state.skipped'
                    $rows += @{
                        Path       = $path
                        Type       = $type
                        Status     = $word.Th
                        StatusEn   = $word.En
                        _RowStatus = 'INFO'
                    }
                    continue
                }

                $dsOk          = $true
                $dsUnsupported = $false
                $dsMessage     = ''
                foreach ($machineName in $machineNames) {
                    $opPath  = 'data/items' + $path + '/machines/' + $machineName + '/validate'
                    $machOk  = $false
                    $machMsg = ''
                    try {
                        $v = Invoke-PMArcGISAdmin -Root $session.Root -Path $opPath -Token $session.Token `
                                                   -TimeoutSec $session.TimeoutSec -Method Post -AllowStatusError
                        $health = ''
                        if ($v.PSObject.Properties['datastore.overallhealth']) { $health = [string]$v.'datastore.overallhealth' }
                        if ([string]$v.status -eq 'success' -and $health -eq 'Healthy') {
                            $machOk = $true
                        }
                        elseif ($v.PSObject.Properties['messages'] -and $v.messages) { $machMsg = (@($v.messages) -join '; ') }
                        elseif (-not [string]::IsNullOrWhiteSpace($health)) { $machMsg = "machine $machineName reported health: $health" }
                        else { $machMsg = "machine $machineName - no health detail provided by the site" }
                    }
                    catch { $machMsg = $_.Exception.Message }

                    if (-not $machOk) {
                        $dsOk = $false
                        # Some Data Store item sub-types (confirmed: a RabbitMQ-
                        # backed "queue" nosql item) do not expose this
                        # operation at all - the site answers with its standard
                        # "no such resource" text (whether that arrives as a
                        # thrown error or as a non-throwing {"status":"error"}
                        # body -AllowStatusError let through) rather than a
                        # real health result. That is the same class of gap as
                        # validateDataItem's NPE above, not a real outage, so
                        # it is skipped rather than reported as CRIT.
                        if ($machMsg -like '*Could not find resource or operation*') { $dsUnsupported = $true }
                        if ([string]::IsNullOrWhiteSpace($dsMessage)) { $dsMessage = $machMsg }
                    }
                }

                if ($dsUnsupported) {
                    $skipped++
                    $findings += New-PMFinding -Severity 'INFO' -TextKey 'agsdata.finding.skipped' -Values @($path, $type)
                    $word = Get-PMWord -Key 'agsdata.state.skipped'
                    $rows += @{ Path = $path; Type = $type; Status = $word.Th; StatusEn = $word.En; _RowStatus = 'INFO' }
                }
                elseif ($dsOk) {
                    $word = Get-PMWord -Key 'agsdata.state.ok'
                    $rows += @{ Path = $path; Type = $type; Status = $word.Th; StatusEn = $word.En; _RowStatus = 'OK' }
                }
                else {
                    $failed++
                    $findings += New-PMFinding -Severity 'CRIT' -TextKey 'agsdata.finding.failed' -Values @($path, $dsMessage)
                    $word = Get-PMWord -Key 'agsdata.state.failed'
                    $rows += @{ Path = $path; Type = $type; Status = $word.Th; StatusEn = $word.En; _RowStatus = 'CRIT' }
                }
                continue
            }

            # Only Path, Type and Info travel into the validate request body -
            # Info never leaves this loop iteration.
            $itemJson = [pscustomobject]@{ path = $item.path; type = $item.type; info = $item.info } |
                        ConvertTo-Json -Depth 8 -Compress

            $status  = 'WARN'
            $message = ''
            try {
                $v = Invoke-PMArcGISAdmin -Root $session.Root -Path 'data/validateDataItem' -Token $session.Token `
                                           -TimeoutSec $session.TimeoutSec -Method Post -Parameters @{ item = $itemJson } `
                                           -AllowStatusError
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
                Skipped = $skipped
                Items   = @($rows | ForEach-Object { [pscustomobject]@{ Path = $_.Path; Type = $_.Type; Status = $_._RowStatus } })
            })
    }
    finally {
        Restore-PMArcGISCertificatePolicy
    }
}

Register-PMCheck -Id 'AGSDATA' -TitleKey 'agsdata.title' -Function 'Invoke-PMCheckArcGISData'
