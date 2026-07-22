# PMtools check - Network adapters and IP configuration.  ASCII-only; text from i18n.json.
#
# Reads through System.Net.NetworkInformation (the IP Helper API) rather than
# Get-NetAdapter / Get-NetIPConfiguration, deliberately:
#
#   * Get-NetAdapter queries the MSFT_NetAdapter WMI provider, which calls into
#     the NIC driver. On a server with NIC teaming the LBFO provider walks every
#     member adapter, and that query can briefly block the network stack - the
#     reported symptom was RDP freezing and latency spiking mid-assessment.
#   * Get-NetIPConfiguration additionally consults the Network List Service,
#     which can trigger a network profile re-detection.
#
# The IP Helper API is a local, cached read that touches neither. It is also
# about nineteen times faster (140 ms against 2.7 s measured), and it works on
# Windows Server 2008 R2, where the Net* cmdlets do not exist at all.

function Invoke-PMCheckNetwork {

    $columns = @(
        (New-PMColumn -Key 'Name'    -TextKey 'net.col.name'),
        (New-PMColumn -Key 'State'   -TextKey 'net.col.state'),
        (New-PMColumn -Key 'Speed'   -TextKey 'net.col.speed' -Align 'right'),
        (New-PMColumn -Key 'Ipv4'    -TextKey 'net.col.ipv4'),
        (New-PMColumn -Key 'Gateway' -TextKey 'net.col.gateway'),
        (New-PMColumn -Key 'Dns'     -TextKey 'net.col.dns' -Wide),
        (New-PMColumn -Key 'Mac'     -TextKey 'net.col.mac')
    )

    $adapters = @([System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
        Where-Object {
            $_.OperationalStatus -eq 'Up' -and
            $_.NetworkInterfaceType -ne 'Loopback' -and
            $_.NetworkInterfaceType -ne 'Tunnel'
        })

    $rows       = @()
    $raw        = @()
    $anyGateway = $false
    $anyDns     = $false

    foreach ($a in ($adapters | Sort-Object Name)) {

        $props = $a.GetIPProperties()

        $ipv4 = @($props.UnicastAddresses |
                  Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' } |
                  ForEach-Object { "$($_.Address)/$($_.PrefixLength)" })

        $gw = @($props.GatewayAddresses |
                Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' -and $_.Address.ToString() -ne '0.0.0.0' } |
                ForEach-Object { $_.Address.ToString() })

        $dns = @($props.DnsAddresses |
                 Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
                 ForEach-Object { $_.ToString() })

        if ($gw.Count  -gt 0) { $anyGateway = $true }
        if ($dns.Count -gt 0) { $anyDns     = $true }

        # Speed is bits per second; network rates are decimal, not binary.
        if ($a.Speed -ge 1000000000) { $speed = ("{0:N1} Gbps" -f ($a.Speed / 1e9)) }
        elseif ($a.Speed -gt 0)      { $speed = ("{0:N0} Mbps" -f ($a.Speed / 1e6)) }
        else                         { $speed = '' }

        $mac = $a.GetPhysicalAddress().ToString()
        if ($mac.Length -eq 12) { $mac = ($mac -replace '(..)(?=.)', '$1-') }

        # An adapter with no IPv4 address is reported but not judged: virtual
        # switches and standby team members legitimately look like this.
        if ($ipv4.Count -gt 0) { $status = 'OK' } else { $status = 'INFO' }

        $rows += New-PMRow -Status $status -Values @{
            Name    = $a.Name
            State   = [string]$a.OperationalStatus
            Speed   = $speed
            Ipv4    = ($ipv4 -join ', ')
            Gateway = ($gw -join ', ')
            Dns     = ($dns -join ', ')
            Mac     = $mac
        }

        $raw += [pscustomobject]@{
            Name = $a.Name; Description = $a.Description; Type = [string]$a.NetworkInterfaceType
            Status = [string]$a.OperationalStatus; SpeedBps = $a.Speed; MacAddress = $mac
            IPv4 = $ipv4; Gateway = $gw; DnsServers = $dns
        }
    }

    # Judged for the machine as a whole, not per adapter. A server with virtual
    # switches, host-only networks or standby team members has several adapters
    # that correctly have no gateway and no DNS; warning about each of those
    # would bury the one condition that actually matters - the machine having
    # no route out, or no way to resolve a name, at all.
    $findings = @()
    if ($rows.Count -gt 0) {
        if (-not $anyGateway) { $findings += New-PMFinding -Severity 'WARN' -TextKey 'net.finding.nogw.all' }
        if (-not $anyDns)     { $findings += New-PMFinding -Severity 'WARN' -TextKey 'net.finding.nodns.all' }
    }

    if ($findings.Count -gt 0) { $status = 'WARN'; $sumKey = 'net.summary.issue' }
    else                       { $status = 'OK';   $sumKey = 'net.summary' }

    return New-PMResult -Id 'NET' -TitleKey 'net.title' -Status $status `
        -SummaryKey $sumKey -SummaryValues @($rows.Count) `
        -Columns $columns -Rows $rows -Findings $findings -Raw $raw
}

Register-PMCheck -Id 'NET' -TitleKey 'net.title' -Function 'Invoke-PMCheckNetwork'
