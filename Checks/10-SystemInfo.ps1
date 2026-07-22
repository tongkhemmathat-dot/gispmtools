# PMtools check - System information and uptime.  ASCII-only; text comes from i18n.json.

function Invoke-PMCheckSystemInfo {

    $os   = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $cs   = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
    $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
    $cpu  = @(Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue)

    $lastBoot = $os.LastBootUpTime
    $span     = (Get-Date) - $lastBoot
    $upDays   = [math]::Floor($span.TotalDays)

    $upStatus = Test-PMThreshold -Name 'UptimeDays' -Value $span.TotalDays
    if ($upStatus -eq 'CRIT') { $upStatus = 'WARN' }   # long uptime is never a CRIT on its own

    $uptimeText = (Get-PMText -Key 'sys.uptimeValue' -Values @($upDays, $span.Hours, $span.Minutes)).Th

    if ($cs -and $cs.PartOfDomain) { $domain = $cs.Domain } elseif ($cs) { $domain = $cs.Workgroup } else { $domain = '' }

    $coreCount = ($cpu | Measure-Object -Property NumberOfCores -Sum).Sum
    $logCount  = ($cpu | Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
    if ($cpu.Count -gt 0) { $cpuName = $cpu[0].Name } else { $cpuName = '' }

    $rows = @(
        (New-PMItemRow -TextKey 'sys.item.hostname'     -Value $env:COMPUTERNAME)
        (New-PMItemRow -TextKey 'sys.item.os'           -Value $os.Caption)
        (New-PMItemRow -TextKey 'sys.item.osversion'    -Value ("{0} (Build {1})" -f $os.Version, $os.BuildNumber))
        (New-PMItemRow -TextKey 'sys.item.arch'         -Value $os.OSArchitecture)
        (New-PMItemRow -TextKey 'sys.item.manufacturer' -Value $(if ($cs) { $cs.Manufacturer } else { '' }))
        (New-PMItemRow -TextKey 'sys.item.model'        -Value $(if ($cs) { $cs.Model } else { '' }))
        (New-PMItemRow -TextKey 'sys.item.serial'       -Value $(if ($bios) { $bios.SerialNumber } else { '' }))
        (New-PMItemRow -TextKey 'sys.item.bios'         -Value $(if ($bios) { ("{0} {1}" -f $bios.Manufacturer, $bios.SMBIOSBIOSVersion) } else { '' }))
        (New-PMItemRow -TextKey 'sys.item.cpu'          -Value $cpuName)
        (New-PMItemRow -TextKey 'sys.item.cores'        -Value ("{0} / {1}" -f $coreCount, $logCount))
        (New-PMItemRow -TextKey 'sys.item.ram'          -Value ("{0} GB" -f (ConvertTo-PMGB ($os.TotalVisibleMemorySize * 1KB) 1)))
        (New-PMItemRow -TextKey 'sys.item.installdate'  -Value (Format-PMStamp $os.InstallDate))
        (New-PMItemRow -TextKey 'sys.item.lastboot'     -Value (Format-PMStamp $lastBoot))
        (New-PMItemRow -TextKey 'sys.item.uptime'       -Value $uptimeText -Status $upStatus)
        (New-PMItemRow -TextKey 'sys.item.domain'       -Value $domain)
        (New-PMItemRow -TextKey 'sys.item.timezone'     -Value ([System.TimeZoneInfo]::Local.DisplayName))
    )

    $findings = @()
    if ($upStatus -eq 'WARN') {
        $findings += New-PMFinding -Severity 'WARN' -TextKey 'sys.finding.uptime' `
            -Values @($upDays, (Get-PMThresholdValue -Name 'UptimeDays' -Level 'Warn'))
    }

    return New-PMResult -Id 'SYSTEM' -TitleKey 'sys.title' `
        -Status (Get-PMWorstStatus @($upStatus, 'INFO')) `
        -SummaryKey 'sys.summary' -SummaryValues @($os.Caption, $upDays) `
        -Columns (New-PMItemColumns) -Rows $rows -Findings $findings `
        -Raw ([pscustomobject]@{
            Caption      = $os.Caption
            Version      = $os.Version
            Build        = $os.BuildNumber
            Architecture = $os.OSArchitecture
            InstallDate  = $os.InstallDate
            LastBootTime = $lastBoot
            UptimeDays   = [math]::Round($span.TotalDays, 2)
            Manufacturer = $(if ($cs) { $cs.Manufacturer } else { $null })
            Model        = $(if ($cs) { $cs.Model } else { $null })
            SerialNumber = $(if ($bios) { $bios.SerialNumber } else { $null })
            Processor    = $cpuName
            Cores        = $coreCount
            LogicalCores = $logCount
            TotalMemoryGB = (ConvertTo-PMGB ($os.TotalVisibleMemorySize * 1KB) 2)
            Domain       = $domain
        })
}

Register-PMCheck -Id 'SYSTEM' -TitleKey 'sys.title' -Function 'Invoke-PMCheckSystemInfo'
