# PMtools check - ArcGIS Server site and machine status.
# ASCII-only; text comes from i18n.json.
#
# Needs a configured connection (launcher -> ArcGIS Server connection) and
# reaches the site over HTTPS, so it is disabled by default in settings.json
# for the same reason WU and CONN are: it leaves the machine.
#
# Read-only: /info and /machines are both plain GETs.

function Invoke-PMCheckArcGISSite {

    try {
        $session = Get-PMArcGISSession

        $info = Invoke-PMArcGISAdmin -Root $session.Root -Path 'info' -Token $session.Token -TimeoutSec $session.TimeoutSec

        $columns  = New-PMItemColumns
        $rows     = @()
        $findings = @()

        $rows += New-PMItemRow -TextKey 'ags.item.url'     -Value $session.Root
        $rows += New-PMItemRow -TextKey 'ags.item.user'    -Value $session.Username
        if ($info.currentversion) { $rows += New-PMItemRow -TextKey 'ags.item.version' -Value ([string]$info.currentversion) }
        if ($info.fullVersion)    { $rows += New-PMItemRow -TextKey 'ags.item.build'   -Value ([string]$info.fullVersion) }

        # An account may be allowed to sign in and still not be allowed to
        # list machines. That is a privilege problem worth reporting, not a
        # reason to fail the whole check - the version above is already
        # useful on its own.
        $machines    = @()
        $machineError = ''
        try {
            $resp = Invoke-PMArcGISAdmin -Root $session.Root -Path 'machines' -Token $session.Token -TimeoutSec $session.TimeoutSec
            if ($resp.machines) { $machines = @($resp.machines) }
        }
        catch { $machineError = $_.Exception.Message }

        $started = 0
        $stopped = 0
        $raw     = @()

        foreach ($m in $machines) {
            $name  = [string]$m.machineName
            $state = [string]$m.configuredState

            # The site reports what it was told to do (configuredState) and,
            # when it can reach the machine, what is actually happening
            # (realTimeState). They disagree exactly when something is wrong.
            $real = ''
            if ($m.PSObject.Properties['realTimeState']) { $real = [string]$m.realTimeState }

            if ($real) { $effective = $real } else { $effective = $state }

            if ($effective -eq 'STARTED') {
                $started++
                $rowStatus = 'OK'
                $valueKey  = 'ags.machine.started'
            }
            else {
                $stopped++
                $rowStatus = 'CRIT'
                $valueKey  = 'ags.machine.stopped'
            }

            $word = Get-PMWord -Key $valueKey
            $rows += @{
                Item       = $name
                ItemEn     = ''
                Value      = $word.Th
                ValueEn    = $word.En
                _RowStatus = $rowStatus
            }

            if ($rowStatus -eq 'CRIT') {
                $findings += New-PMFinding -Severity 'CRIT' -TextKey 'ags.finding.machineStopped' -Values @($name)
            }

            $raw += [pscustomobject]@{
                MachineName = $name; ConfiguredState = $state; RealTimeState = $real
                Platform = [string]$m.platform; AdminURL = [string]$m.adminURL
            }
        }

        if ($machineError) {
            $findings += New-PMFinding -Severity 'WARN' -TextKey 'ags.finding.noMachineAccess' -Values @($machineError)
        }

        if ($stopped -gt 0) {
            $status = 'CRIT'
            $sumKey = 'ags.summary.issue'
            $sumVal = @($machines.Count, $stopped)
        }
        elseif ($machineError) {
            $status = 'WARN'
            $sumKey = 'ags.summary.partial'
            $sumVal = @([string]$info.currentversion)
        }
        else {
            $status = 'OK'
            $sumKey = 'ags.summary.ok'
            $sumVal = @([string]$info.currentversion, $machines.Count)
        }

        return New-PMResult -Id 'AGS' -TitleKey 'ags.title' -Status $status `
            -SummaryKey $sumKey -SummaryValues $sumVal `
            -Columns $columns -Rows $rows -Findings $findings `
            -Raw ([pscustomobject]@{ Root = $session.Root; Info = $info; Machines = $raw })
    }
    finally {
        # The certificate callback is process-wide with no scoped form on
        # PS 5.1, so it is put back as soon as this check is done rather
        # than left switched off for whatever runs next.
        Restore-PMArcGISCertificatePolicy
    }
}

Register-PMCheck -Id 'AGS' -TitleKey 'ags.title' -Function 'Invoke-PMCheckArcGISSite'
