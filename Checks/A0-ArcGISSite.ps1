# PMtools check - ArcGIS Server site and machine status.
# ASCII-only; text comes from i18n.json.
#
# Needs a configured connection (launcher -> ArcGIS Server connection) and
# reaches the site over HTTPS, so it is disabled by default in settings.json
# for the same reason WU and CONN are: it leaves the machine.
#
# Read-only: /info, /machines, /machines/<name>/status and /about are all
# plain GETs.
#
# /admin/about (confirmed against a real site, ArcGIS Server 11.5.0, on
# 2026-07-23) is not documented anywhere this project had seen before
# testing it directly. One call returns hardware/OS facts for every
# machine in the site at once (machines[].hardware: os, cpu, memory), plus
# license detail this check does not use. localDiskUsage came back an
# empty array on the real site - its populated shape is unverified, so it
# is deliberately left unread rather than guessed at.

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
        $machines     = @()
        $machineError = ''
        try {
            $resp = Invoke-PMArcGISAdmin -Root $session.Root -Path 'machines' -Token $session.Token -TimeoutSec $session.TimeoutSec
            if ($resp.machines) { $machines = @($resp.machines) }
        }
        catch { $machineError = $_.Exception.Message }

        # Hardware/OS detail per machine, all in this one extra call. Kept
        # separate from the /machines + /machines/<name>/status calls above:
        # this is supplementary detail for the report, not something the
        # status roll-up depends on, so a failure here only downgrades to a
        # single finding rather than affecting $status at all.
        $hwByMachine = @{}
        $aboutError  = ''
        try {
            $about = Invoke-PMArcGISAdmin -Root $session.Root -Path 'about' -Token $session.Token -TimeoutSec $session.TimeoutSec
            foreach ($hw in @($about.machines)) {
                $hwName = [string]$hw.machineName
                if (-not [string]::IsNullOrWhiteSpace($hwName)) { $hwByMachine[$hwName] = $hw }
            }
        }
        catch { $aboutError = $_.Exception.Message }

        $started = 0
        $stopped = 0
        $unknown = 0
        $raw     = @()

        foreach ($m in $machines) {
            $name = [string]$m.machineName

            # /admin/machines lists only machineName, adminURL, synchronize
            # and underMaintenance - it carries NO state at all. Verified
            # against ArcGIS Server 11.5; the run state has to be fetched
            # per machine from /machines/<name>/status.
            #
            # The first version of this check read $m.configuredState
            # straight off the list entry, got an empty string for every
            # machine, and reported a healthy two-machine production site as
            # two CRITical outages. That is why an unreadable state is now
            # treated as UNKNOWN rather than folded in with "stopped": a
            # monitoring tool that cries wolf on a healthy site is worse than
            # one that admits it does not know.
            $configured = ''
            $real       = ''
            $stateError = ''
            try {
                $st = Invoke-PMArcGISAdmin -Root $session.Root -Path ("machines/$name/status") `
                                           -Token $session.Token -TimeoutSec $session.TimeoutSec
                if ($st.PSObject.Properties['configuredState']) { $configured = [string]$st.configuredState }
                if ($st.PSObject.Properties['realTimeState'])   { $real       = [string]$st.realTimeState }
            }
            catch { $stateError = $_.Exception.Message }

            # realTimeState is what is actually happening; configuredState is
            # only what the site was told to do. Prefer the former.
            if ($real) { $effective = $real } else { $effective = $configured }

            $maintenance = $false
            if ($m.PSObject.Properties['underMaintenance']) { $maintenance = [bool]$m.underMaintenance }

            if ([string]::IsNullOrWhiteSpace($effective)) {
                $unknown++
                $rowStatus = 'WARN'
                $valueKey  = 'ags.machine.unknown'
                $reason    = $stateError
                if ([string]::IsNullOrWhiteSpace($reason)) { $reason = 'the status resource returned no state' }
                $findings += New-PMFinding -Severity 'WARN' -TextKey 'ags.finding.machineUnknown' -Values @($name, $reason)
            }
            elseif ($effective -eq 'STARTED') {
                $started++
                if ($maintenance) {
                    $rowStatus = 'WARN'
                    $valueKey  = 'ags.machine.maintenance'
                    $findings += New-PMFinding -Severity 'WARN' -TextKey 'ags.finding.machineMaintenance' -Values @($name)
                }
                else {
                    $rowStatus = 'OK'
                    $valueKey  = 'ags.machine.started'
                }
            }
            else {
                $stopped++
                $rowStatus = 'CRIT'
                $valueKey  = 'ags.machine.stopped'
                $findings += New-PMFinding -Severity 'CRIT' -TextKey 'ags.finding.machineStopped' -Values @($name)
            }

            $word = Get-PMWord -Key $valueKey

            # Hardware detail is a proper-noun/technical fact (OS name, CPU
            # model, memory in GB) rather than something to translate, so the
            # same text is appended to both the Thai and English cell - same
            # precedent as raw service Type strings elsewhere in the A*-ArcGIS
            # checks.
            $valueTh = $word.Th
            $valueEn = $word.En
            $hwRaw   = $null
            if ($hwByMachine.ContainsKey($name)) {
                $hw = $hwByMachine[$name].hardware
                if ($hw) {
                    $os       = [string]$hw.os
                    $cpuModel = ''
                    $cpuLines = @([string]$hw.cpu -split "`n")
                    if ($cpuLines.Count -gt 0) { $cpuModel = $cpuLines[0].Trim() }
                    $cores = 0
                    if ($hw.numLogicalProcessors) { $cores = [int]$hw.numLogicalProcessors }
                    $totalGB = 0.0; $freeGB = 0.0
                    if ($hw.systemMemoryMB)          { $totalGB = [math]::Round(([double]$hw.systemMemoryMB) / 1024, 1) }
                    if ($hw.systemMemoryAvailableMB) { $freeGB  = [math]::Round(([double]$hw.systemMemoryAvailableMB) / 1024, 1) }

                    $hwText = ('{0} | {1} ({2} cores) | RAM {3} GB ({4} GB free)' -f $os, $cpuModel, $cores, $totalGB, $freeGB)
                    $valueTh = "$valueTh - $hwText"
                    $valueEn = "$valueEn - $hwText"
                    $hwRaw   = [pscustomobject]@{ Os = $os; Cpu = $cpuModel; LogicalProcessors = $cores; MemoryTotalGB = $totalGB; MemoryFreeGB = $freeGB }
                }
            }

            $rows += @{
                Item       = $name
                ItemEn     = ''
                Value      = $valueTh
                ValueEn    = $valueEn
                _RowStatus = $rowStatus
            }

            $raw += [pscustomobject]@{
                MachineName = $name; ConfiguredState = $configured; RealTimeState = $real
                UnderMaintenance = $maintenance; AdminURL = [string]$m.adminURL
                StateError = $stateError; Hardware = $hwRaw
            }
        }

        if ($aboutError) {
            $findings += New-PMFinding -Severity 'INFO' -TextKey 'ags.finding.aboutError' -Values @($aboutError)
        }

        if ($machineError) {
            $findings += New-PMFinding -Severity 'WARN' -TextKey 'ags.finding.noMachineAccess' -Values @($machineError)
        }

        $version = [string]$info.currentversion

        if ($stopped -gt 0) {
            $status = 'CRIT'
            $sumKey = 'ags.summary.issue'
            $sumVal = @($machines.Count, $stopped)
        }
        elseif ($unknown -gt 0) {
            $status = 'WARN'
            $sumKey = 'ags.summary.unknown'
            $sumVal = @($version, $machines.Count, $unknown)
        }
        elseif ($machineError) {
            $status = 'WARN'
            $sumKey = 'ags.summary.partial'
            $sumVal = @($version)
        }
        else {
            # A machine in maintenance still counts as reachable and healthy
            # for the roll-up; its own row and finding carry the warning.
            $status = Get-PMWorstStatus (@($rows | ForEach-Object { $_._RowStatus }) + @('OK'))
            $sumKey = 'ags.summary.ok'
            $sumVal = @($version, $machines.Count)
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
