# PMtools check - Reachability of gateways, DNS servers and any extra targets.
# ASCII-only; text comes from i18n.json.
#
# DISABLED BY DEFAULT in Config\settings.json (Checks.Disabled).
#
# This is the only check that puts packets on the wire. On a site with an
# IPS/IDS, repeated ICMP from a server to its gateway and every DNS server can
# be scored as a scan and get the host throttled or blocked - which looks
# exactly like the network going unstable during the assessment. Everything
# else PMtools does is a local read.
#
# Enable it when the network team is happy with it:
#     .\Start-PMCheck.ps1 -Only CONN
# or remove "CONN" from Checks.Disabled in settings.json.
#
# A failure here is reported as WARN rather than CRIT on purpose: plenty of
# hardened gateways and DNS servers are configured not to answer ICMP at all,
# so an unreachable result is a prompt to look, not proof of a fault. The
# finding text says so.

function Invoke-PMCheckConnectivity {

    $targets = @()

    $tGw  = Get-PMWord -Key 'conn.type.gateway'
    $tDns = Get-PMWord -Key 'conn.type.dns'
    $tCus = Get-PMWord -Key 'conn.type.custom'

    # Targets come from the IP Helper API for the same reason the NET check uses
    # it: Get-NetIPConfiguration and Get-DnsClientServerAddress go through the
    # NIC driver's WMI provider and the Network List Service, which can stall
    # the network stack on a teamed server.
    $interfaces = @([System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
        Where-Object { $_.OperationalStatus -eq 'Up' -and $_.NetworkInterfaceType -ne 'Loopback' })

    if (Get-PMSetting -Path 'Connectivity.PingGateway' -Default $true) {
        $gws = @($interfaces | ForEach-Object { $_.GetIPProperties().GatewayAddresses } |
                 Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' -and $_.Address.ToString() -ne '0.0.0.0' } |
                 ForEach-Object { $_.Address.ToString() })
        foreach ($g in @($gws | Select-Object -Unique)) {
            $targets += [pscustomobject]@{ Address = $g; TypeTh = $tGw.Th; TypeEn = $tGw.En }
        }
    }

    if (Get-PMSetting -Path 'Connectivity.PingDns' -Default $true) {
        $dnsList = @($interfaces | ForEach-Object { $_.GetIPProperties().DnsAddresses } |
                     Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
                     ForEach-Object { $_.ToString() })
        foreach ($d in @($dnsList | Where-Object { $_ -ne '127.0.0.1' } | Select-Object -Unique)) {
            $targets += [pscustomobject]@{ Address = $d; TypeTh = $tDns.Th; TypeEn = $tDns.En }
        }
    }

    foreach ($e in @(Get-PMSetting -Path 'Connectivity.ExtraTargets' -Default @())) {
        if ($e) { $targets += [pscustomobject]@{ Address = $e; TypeTh = $tCus.Th; TypeEn = $tCus.En } }
    }

    $columns = @(
        (New-PMColumn -Key 'Target'  -TextKey 'conn.col.target'),
        (New-PMColumn -Key 'Type'    -TextKey 'conn.col.type'),
        (New-PMColumn -Key 'Result'  -TextKey 'conn.col.result'),
        (New-PMColumn -Key 'Latency' -TextKey 'conn.col.latency' -Align 'right')
    )

    if ($targets.Count -eq 0) {
        return New-PMResult -Id 'CONN' -TitleKey 'conn.title' -Status 'INFO' `
            -SummaryKey 'conn.summary.none' `
            -Columns $columns -Rows @() -Findings @() `
            -Raw ([pscustomobject]@{ TargetCount = 0 })
    }

    $okWord   = Get-PMWord -Key 'conn.result.ok'
    $failWord = Get-PMWord -Key 'conn.result.fail'

    $rows     = @()
    $findings = @()
    $raw      = @()
    $failed   = 0

    # Test-Connection on PowerShell 5.1 has no timeout parameter and waits out
    # the full default on every unreachable address, which turned this check
    # into a 50-second stall. System.Net.NetworkInformation.Ping takes an
    # explicit timeout, capping the whole check at roughly 2 s per target.
    $pinger  = New-Object System.Net.NetworkInformation.Ping
    $timeout = 1000

    foreach ($t in ($targets | Sort-Object Address -Unique)) {

        $times = @()
        for ($attempt = 0; $attempt -lt 2; $attempt++) {
            try {
                $reply = $pinger.Send($t.Address, $timeout)
                if ($reply -and $reply.Status -eq 'Success') { $times += [double]$reply.RoundtripTime }
            }
            catch { }   # unresolvable name or no route: counts as unreachable
        }

        if ($times.Count -gt 0) {
            $latency = [math]::Round(($times | Measure-Object -Average).Average, 0)
            $status  = 'OK'; $resTh = $okWord.Th; $resEn = $okWord.En
        }
        else {
            $latency = '-'
            $status  = 'WARN'; $resTh = $failWord.Th; $resEn = $failWord.En
            $failed++
            $findings += New-PMFinding -Severity 'WARN' -TextKey 'conn.finding.fail' -Values @($t.Address, $t.TypeTh)
        }

        $row = New-PMRow -Status $status -Values @{
            Target  = $t.Address
            Type    = $t.TypeTh
            Result  = $resTh
            Latency = $latency
        }
        $row['TypeEn']   = $t.TypeEn
        $row['ResultEn'] = $resEn
        $rows += $row

        $raw += [pscustomobject]@{ Target = $t.Address; Type = $t.TypeEn; Reachable = ($times.Count -gt 0); LatencyMs = $latency }
    }
    $pinger.Dispose()

    if ($failed -gt 0) { $status = 'WARN'; $sumKey = 'conn.summary.issue'; $sumVal = @($rows.Count, $failed) }
    else               { $status = 'OK';   $sumKey = 'conn.summary.ok';    $sumVal = @($rows.Count) }

    return New-PMResult -Id 'CONN' -TitleKey 'conn.title' -Status $status `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns $columns -Rows $rows -Findings $findings -Raw $raw
}

Register-PMCheck -Id 'CONN' -TitleKey 'conn.title' -Function 'Invoke-PMCheckConnectivity'
