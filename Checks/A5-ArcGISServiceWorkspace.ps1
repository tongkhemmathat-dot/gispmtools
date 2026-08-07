# PMtools check - what data source ("service workspace" in ArcGIS Server
# Manager's own terms) each map service actually connects to: server/IP,
# instance, database and user account for an enterprise geodatabase, or a
# path for a file-based connection - broken into separate fields so they
# can be compared at a glance, plus whether that connection matches a data
# store registered on the site. Also classifies each connection as
# Referenced (same data the publisher used), Replaced (a different
# machine/path substituted at publish time) or Copied (data copied onto
# the server at publish time, no longer tied to the source).
# ASCII-only; text comes from i18n.json.
#
# This is the automated form of ArcGIS Server Manager's own "Service
# Workspaces" page (Services > Manage Services > select a service >
# Service Workspaces), which shows exactly these three categories -
# reviewing every service there one at a time does not scale past a
# handful of services.
#
# The "Matching Registered Data Store" column is a NAME LOOKUP only - it
# answers "does this connection correspond to a data store the site has
# registered, and if so which one" using the same key-matching functions
# AGSDATA's peers use (Lib\ArcGIS.ps1: Get-PMArcGISDatabaseKey,
# Get-PMArcGISPathKey, Test-PMArcGISPathMatch). It deliberately stops
# there: no BREAK/LOW/ORPHAN classification, no computeRefCount
# cross-check, no per-store roll-up table - that fuller "what breaks if
# this store is unregistered" analysis was tried once (AGSIMPACT) and
# removed by request as more than was needed. If that is what is wanted
# again later, see docs/HANDOVER.md for what AGSIMPACT did.
#
# Needs a configured connection and is disabled by default, like every
# other A*-ArcGIS* check - see A0-ArcGISSite.ps1.
#
# Read-only: data/findItems is POST because the Admin API requires it for
# that operation (see Lib\ArcGIS.ps1's file header) - it does not write.
# The manifest resource, services/{folder} and the plain service resource
# (the ImageServer/mosaic fallback) are all GET.
#
# The manifest resource's path is
# services/[<folder>/]<name>.<type>/iteminfo/manifest/manifest.json - NOT
# services/[<folder>/]<name>.<type>/manifest, which was tried first and
# confirmed against a real site (2026-08-07) to answer "Could not find
# resource or operation 'manifest' on the system" for every service,
# including the built-in SampleWorldCities - a wrong path, not a wrong
# HTTP method (an earlier fix here switched this call to POST, going by a
# secondhand claim rather than Esri's own reference; that claim was wrong
# and is why this had to be fixed twice - the path was the actual bug
# both times, confirmed this time against Esri's official documentation:
# https://developers.arcgis.com/rest/enterprise-administration/server/servicemanifest.htm).
#
# Runs against every service the site has, not a capped top-N like
# AGSSVC - this check exists specifically to be reviewed as a directory,
# so a service missing from it because it did not make a "busiest N" cut
# would defeat the point.
#
# SECURITY: same rule as AGSDATA - findItems returns each enterprise
# geodatabase item's info.connectionString, which carries an
# ENCRYPTED_PASSWORD blob; a service's own manifest can carry the same in
# its connection strings. Both are parsed in memory only, to read
# SERVER/INSTANCE/DATABASE/USER and to build the match key - the raw
# connection string itself is never placed in $rows, $findings or -Raw.
#
# Field shapes this check relies on - manifest.databases[]
# (onServerConnectionString, onPremiseConnectionString, byReference),
# manifest.resources[] (onPremisePath, serverPath) and findItems items'
# info.connectionString/info.path - come from Esri's published Admin API
# reference. The manifest resource's own shape is now confirmed against a
# real site (2026-08-07); findItems' item shape is not yet.

function Invoke-PMCheckArcGISServiceWorkspace {

    try {
        $session = Get-PMArcGISSession

        # --- registered data stores, for the name-lookup column only -------
        # Only egdb (connection-string based) and folder/rasterStore
        # (path-based) are matchable - cloudStore/nosql/bigDataFileShare are
        # not a plain connection string or filesystem path this check can
        # compare, same limitation AGSDATA's peers documented before.
        $stores   = @()
        $findings = @()
        try {
            $findResp = Invoke-PMArcGISAdmin -Root $session.Root -Path 'data/findItems' -Token $session.Token `
                                              -TimeoutSec $session.TimeoutSec -Method Post `
                                              -Parameters @{ ancestorPath = '/'; types = 'egdb,folder,rasterStore' }
            foreach ($item in @($findResp.items)) {
                $itemPath = [string]$item.path
                $itemType = [string]$item.type

                if ($itemType -eq 'egdb') {
                    if ($item.info -and -not [string]::IsNullOrWhiteSpace([string]$item.info.connectionString)) {
                        $sParts = ConvertFrom-PMArcGISConnectionString -ConnectionString ([string]$item.info.connectionString)
                        $key = Get-PMArcGISDatabaseKey -ConnectionParts $sParts
                        $stores += [pscustomobject]@{ Path = $itemPath; Key = $key; MatchType = 'db'; Readable = $true }
                    }
                }
                elseif ($itemType -eq 'folder' -or $itemType -eq 'rasterStore') {
                    $physPath = ''
                    if ($item.info -and -not [string]::IsNullOrWhiteSpace([string]$item.info.path)) { $physPath = [string]$item.info.path }
                    $key = Get-PMArcGISPathKey -Path $physPath
                    $stores += [pscustomobject]@{ Path = $itemPath; Key = $key; MatchType = 'path'; Readable = (-not [string]::IsNullOrWhiteSpace($key)) }
                }
            }
        }
        catch {
            $findings += New-PMFinding -Severity 'WARN' -TextKey 'agsworkspace.finding.storesError' -Values @($_.Exception.Message)
        }

        function Find-PMWorkspaceStore {
            param([string]$MatchType, [string]$Key)
            if ([string]::IsNullOrWhiteSpace($Key)) { return $null }
            if ($MatchType -eq 'db') {
                return ($stores | Where-Object { $_.MatchType -eq 'db' -and $_.Readable -and $_.Key -eq $Key } | Select-Object -First 1)
            }
            return ($stores | Where-Object { $_.MatchType -eq 'path' -and $_.Readable -and (Test-PMArcGISPathMatch -ServiceKey $Key -StoreKey $_.Key) } | Select-Object -First 1)
        }

        # --- enumerate services ---------------------------------------------
        # Same folder/type filter as the other manifest-reading ArcGIS
        # checks: skip the System/Utilities folders ArcGIS ships with dozens
        # of internal services in, and the two service types that never
        # carry a data source of their own.
        $skipFolders = @('System', 'Utilities')
        $skipTypes   = @('GeometryServer', 'SearchServer')

        $root    = Invoke-PMArcGISAdmin -Root $session.Root -Path 'services' -Token $session.Token -TimeoutSec $session.TimeoutSec
        $folders = @('') + @($root.folders)

        $svcList = @()

        foreach ($folder in $folders) {
            if ($skipFolders -contains $folder) { continue }

            if ([string]::IsNullOrWhiteSpace($folder)) {
                $list = @($root.services)
            }
            else {
                try {
                    $resp = Invoke-PMArcGISAdmin -Root $session.Root -Path "services/$folder" -Token $session.Token -TimeoutSec $session.TimeoutSec
                    $list = @($resp.services)
                }
                catch {
                    $findings += New-PMFinding -Severity 'WARN' -TextKey 'agsworkspace.finding.folderError' -Values @($folder, $_.Exception.Message)
                    continue
                }
            }

            foreach ($s in $list) {
                $type = [string]$s.type
                if ($skipTypes -contains $type) { continue }
                $svcList += [pscustomobject]@{ Folder = $folder; Name = [string]$s.serviceName; Type = $type }
            }
        }

        if ($svcList.Count -eq 0) {
            return New-PMResult -Id 'AGSWORKSPACE' -TitleKey 'agsworkspace.title' -Status 'INFO' `
                -SummaryKey 'agsworkspace.summary.none' -Findings $findings
        }

        $columns = @(
            (New-PMColumn -Key 'Service'         -TextKey 'agsworkspace.col.service' -Wide),
            (New-PMColumn -Key 'Type'            -TextKey 'agsworkspace.col.type'),
            (New-PMColumn -Key 'Category'        -TextKey 'agsworkspace.col.category'),
            (New-PMColumn -Key 'Server'          -TextKey 'agsworkspace.col.server' -Wide),
            (New-PMColumn -Key 'Instance'        -TextKey 'agsworkspace.col.instance'),
            (New-PMColumn -Key 'Database'        -TextKey 'agsworkspace.col.database'),
            (New-PMColumn -Key 'User'            -TextKey 'agsworkspace.col.user'),
            (New-PMColumn -Key 'RegisteredStore' -TextKey 'agsworkspace.col.registeredstore' -Wide),
            (New-PMColumn -Key 'Status'          -TextKey 'agsworkspace.col.status')
        )

        $rows         = @()
        $unknownCount = 0
        $noDataCount  = 0

        foreach ($svc in $svcList) {
            $folder = $svc.Folder; $name = $svc.Name; $type = $svc.Type
            $label  = if ($folder) { "$folder/$name" } else { $name }
            $svcRel = if ($folder) { "$folder/$name.$type" } else { "$name.$type" }

            $manifest      = $null
            $manifestError = ''
            try {
                $manifest = Invoke-PMArcGISAdmin -Root $session.Root -Path "services/$svcRel/iteminfo/manifest/manifest.json" -Token $session.Token -TimeoutSec $session.TimeoutSec
            }
            catch { $manifestError = $_.Exception.Message }

            if ($manifestError) {
                $unknownCount++
                $findings += New-PMFinding -Severity 'WARN' -TextKey 'agsworkspace.finding.unknown' -Values @($label, $manifestError)
                $word = Get-PMWord -Key 'agsworkspace.state.unknown'
                $rows += @{
                    Service = $label; Type = $type; Category = 'Unknown'
                    Server = '?'; Instance = '?'; Database = '?'; User = '?'; RegisteredStore = '?'
                    Status = $word.Th; StatusEn = $word.En; _RowStatus = 'WARN'
                }
                continue
            }

            $workspaces = @()

            foreach ($dbEntry in @($manifest.databases)) {
                $onServerStr = [string]$dbEntry.onServerConnectionString
                if ([string]::IsNullOrWhiteSpace($onServerStr)) { continue }

                $serverParts = ConvertFrom-PMArcGISConnectionString -ConnectionString $onServerStr
                $server   = Get-PMArcGISNormalizedServer -ConnectionParts $serverParts
                $instance = [string]$serverParts['INSTANCE']
                $db       = [string]$serverParts['DATABASE']
                $user     = [string]$serverParts['USER']
                $key      = Get-PMArcGISDatabaseKey -ConnectionParts $serverParts
                $store    = Find-PMWorkspaceStore -MatchType 'db' -Key $key

                $byRef = $false
                if ($dbEntry.PSObject.Properties['byReference']) { $byRef = [bool]$dbEntry.byReference }

                if (-not $byRef) {
                    $category = 'Copied'
                }
                else {
                    $onPremStr = [string]$dbEntry.onPremiseConnectionString
                    if ([string]::IsNullOrWhiteSpace($onPremStr)) {
                        # Nothing to compare the server-side connection
                        # against - the common single-machine case. Assume
                        # Referenced rather than guess Replaced.
                        $category = 'Referenced'
                    }
                    else {
                        $premParts  = ConvertFrom-PMArcGISConnectionString -ConnectionString $onPremStr
                        $premServer = Get-PMArcGISNormalizedServer -ConnectionParts $premParts
                        $premDb     = [string]$premParts['DATABASE']
                        $premUser   = [string]$premParts['USER']
                        if ($premServer -eq $server -and $premDb -eq $db -and $premUser -eq $user) { $category = 'Referenced' }
                        else { $category = 'Replaced' }
                    }
                }

                $workspaces += [pscustomobject]@{
                    Category = $category; Server = $server; Instance = $instance; Database = $db; User = $user
                    RegisteredStore = $(if ($store) { $store.Path } else { '-' })
                }
            }

            foreach ($res in @($manifest.resources)) {
                $onPrem     = [string]$res.onPremisePath
                $serverPath = [string]$res.serverPath
                $shown      = if (-not [string]::IsNullOrWhiteSpace($serverPath)) { $serverPath } else { $onPrem }
                # /vsi... is a GDAL virtual-filesystem prefix ArcGIS uses for
                # a cloud-store-backed raster path - not a Windows/UNC path
                # this check tries to classify Referenced/Replaced for.
                #
                # manifest.resources[] also always lists the service's OWN
                # map document (.msd/.mxd/.sd), bundled into the server's
                # internal arcgissystem\arcgisinput\...\extracted\... folder
                # at publish time - confirmed against a real site
                # (2026-08-08): every service in a 10-service sample carried
                # exactly one such entry. That is an implementation artifact,
                # not a data workspace to compare - skipped so it does not
                # clutter the Server/Instance/Database/User columns next to
                # the service's real database or file connection.
                if ([string]::IsNullOrWhiteSpace($shown) -or $shown -match '^/vsi') { continue }
                if ($shown -match '\.(msd|mxd|sd)$' -or $shown -match '\\arcgissystem\\') { continue }

                if (-not [string]::IsNullOrWhiteSpace($onPrem) -and -not [string]::IsNullOrWhiteSpace($serverPath)) {
                    $premKey = $onPrem.Trim().Replace('/', '\').TrimEnd('\').ToLower()
                    $srvKey  = $serverPath.Trim().Replace('/', '\').TrimEnd('\').ToLower()
                    $category = if ($premKey -eq $srvKey) { 'Referenced' } else { 'Replaced' }
                }
                else {
                    $category = 'Referenced'
                }

                $key   = Get-PMArcGISPathKey -Path $shown
                $store = Find-PMWorkspaceStore -MatchType 'path' -Key $key

                $workspaces += [pscustomobject]@{
                    Category = $category; Server = $shown; Instance = ''; Database = ''; User = ''
                    RegisteredStore = $(if ($store) { $store.Path } else { '-' })
                }
            }

            if ($workspaces.Count -eq 0) {
                # manifest.databases and manifest.resources were both empty -
                # ImageServer/mosaic dataset connections in particular live
                # in the service's own properties instead of the manifest.
                # No Referenced/Replaced/Copied distinction is available
                # from this fallback, so it is always shown as Referenced.
                try {
                    $propResp = Invoke-PMArcGISAdmin -Root $session.Root -Path "services/$svcRel" -Token $session.Token -TimeoutSec $session.TimeoutSec
                    foreach ($propName in @('path', 'filePath', 'locatorPath', 'cacheDir')) {
                        if ($propResp.properties -and $propResp.properties.PSObject.Properties[$propName]) {
                            $v = [string]$propResp.properties.$propName
                            if (-not [string]::IsNullOrWhiteSpace($v) -and $v -notmatch '^/vsi') {
                                $key   = Get-PMArcGISPathKey -Path $v
                                $store = Find-PMWorkspaceStore -MatchType 'path' -Key $key
                                $workspaces += [pscustomobject]@{
                                    Category = 'Referenced'; Server = $v; Instance = ''; Database = ''; User = ''
                                    RegisteredStore = $(if ($store) { $store.Path } else { '-' })
                                }
                                break
                            }
                        }
                    }
                }
                catch { }   # best-effort fallback only - leaves this service in the "no data source" bucket, not UNKNOWN
            }

            if ($workspaces.Count -eq 0) {
                $noDataCount++
                $word = Get-PMWord -Key 'agsworkspace.state.ok'
                $rows += @{
                    Service = $label; Type = $type; Category = 'None'
                    Server = '-'; Instance = '-'; Database = '-'; User = '-'; RegisteredStore = '-'
                    Status = $word.Th; StatusEn = $word.En; _RowStatus = 'INFO'
                }
                continue
            }

            # Almost always exactly one entry - joined with "; " (positionally
            # matched across columns) for the rare service with more than one.
            $categoryDisplay = (@($workspaces | ForEach-Object { $_.Category } | Sort-Object -Unique)) -join ', '
            $join = { param($field) (@($workspaces | ForEach-Object { if ([string]::IsNullOrWhiteSpace($_.$field)) { '-' } else { $_.$field } })) -join '; ' }
            $word = Get-PMWord -Key 'agsworkspace.state.ok'
            $rows += @{
                Service = $label; Type = $type; Category = $categoryDisplay
                Server = (& $join 'Server'); Instance = (& $join 'Instance'); Database = (& $join 'Database'); User = (& $join 'User')
                RegisteredStore = (& $join 'RegisteredStore')
                Status = $word.Th; StatusEn = $word.En; _RowStatus = 'OK'
            }
        }

        if ($unknownCount -gt 0) {
            $status = 'WARN'
            $sumKey = 'agsworkspace.summary.issue'
            $sumVal = @($svcList.Count, $unknownCount)
        }
        else {
            $status = 'OK'
            $sumKey = 'agsworkspace.summary.ok'
            $sumVal = @($svcList.Count)
        }

        return New-PMResult -Id 'AGSWORKSPACE' -TitleKey 'agsworkspace.title' -Status $status `
            -SummaryKey $sumKey -SummaryValues $sumVal `
            -Columns $columns -Rows $rows -Findings $findings `
            -Raw ([pscustomobject]@{
                Total          = $svcList.Count
                UnknownCount   = $unknownCount
                NoDataCount    = $noDataCount
                RegisteredStores = $stores.Count
            })
    }
    finally {
        Restore-PMArcGISCertificatePolicy
    }
}

Register-PMCheck -Id 'AGSWORKSPACE' -TitleKey 'agsworkspace.title' -Function 'Invoke-PMCheckArcGISServiceWorkspace'
