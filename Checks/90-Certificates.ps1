# PMtools check - Machine certificate expiry.  ASCII-only; text comes from i18n.json.
# Covers the personal store and, where present, the IIS web hosting store.

function Invoke-PMCheckCertificates {

    $storePaths = @('Cert:\LocalMachine\My', 'Cert:\LocalMachine\WebHosting')

    $certs = @()
    foreach ($path in $storePaths) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $found = @(Get-ChildItem -LiteralPath $path -ErrorAction SilentlyContinue)
        foreach ($c in $found) {
            $certs += [pscustomobject]@{ Cert = $c; Store = ($path -replace '^Cert:\\LocalMachine\\', '') }
        }
    }

    $columns = @(
        (New-PMColumn -Key 'Subject'    -TextKey 'cert.col.subject' -Wide),
        (New-PMColumn -Key 'Issuer'     -TextKey 'cert.col.issuer'),
        (New-PMColumn -Key 'Expiry'     -TextKey 'cert.col.expiry'),
        (New-PMColumn -Key 'DaysLeft'   -TextKey 'cert.col.daysleft' -Align 'right'),
        (New-PMColumn -Key 'Thumbprint' -TextKey 'cert.col.thumbprint')
    )

    if ($certs.Count -eq 0) {
        return New-PMResult -Id 'CERT' -TitleKey 'cert.title' -Status 'INFO' `
            -SummaryKey 'cert.summary.none' `
            -Columns $columns -Rows @() -Findings @() `
            -Raw ([pscustomobject]@{ CertificateCount = 0 })
    }

    $rows     = @()
    $findings = @()
    $raw      = @()
    $issues   = 0

    foreach ($item in ($certs | Sort-Object { $_.Cert.NotAfter })) {

        $c        = $item.Cert
        $daysLeft = [math]::Floor(([datetime]$c.NotAfter - (Get-Date)).TotalDays)
        $name     = $c.Subject
        if ($c.FriendlyName) { $name = $c.FriendlyName }
        $shortName = Get-PMShortText -Text $name -MaxLength 70

        # Test-PMThreshold handles the near-expiry bands; an already-expired
        # certificate is its own case and always critical.
        if ($daysLeft -lt 0) { $status = 'CRIT' }
        else                 { $status = Test-PMThreshold -Name 'CertExpiryDays' -Value $daysLeft }

        if ($status -ne 'OK') { $issues++ }

        $rows += New-PMRow -Status $status -Values @{
            Subject    = Get-PMShortText -Text $c.Subject -MaxLength 80
            Issuer     = Get-PMShortText -Text $c.Issuer  -MaxLength 50
            Expiry     = ([datetime]$c.NotAfter).ToString('yyyy-MM-dd')
            DaysLeft   = $daysLeft
            Thumbprint = $c.Thumbprint
        }

        if ($daysLeft -lt 0) {
            $findings += New-PMFinding -Severity 'CRIT' -TextKey 'cert.finding.expired' `
                -Values @($shortName, ([datetime]$c.NotAfter).ToString('yyyy-MM-dd'))
        }
        elseif ($status -eq 'CRIT') {
            $findings += New-PMFinding -Severity 'CRIT' -TextKey 'cert.finding.crit' `
                -Values @($shortName, $daysLeft, ([datetime]$c.NotAfter).ToString('yyyy-MM-dd'))
        }
        elseif ($status -eq 'WARN') {
            $findings += New-PMFinding -Severity 'WARN' -TextKey 'cert.finding.warn' `
                -Values @($shortName, $daysLeft, ([datetime]$c.NotAfter).ToString('yyyy-MM-dd'))
        }

        $raw += [pscustomobject]@{
            Subject = $c.Subject; Issuer = $c.Issuer; NotAfter = $c.NotAfter
            DaysLeft = $daysLeft; Thumbprint = $c.Thumbprint; Store = $item.Store; Status = $status
        }
    }

    $overall = Get-PMWorstStatus (@($rows | ForEach-Object { $_._RowStatus }) + @('OK'))

    if ($issues -gt 0) { $sumKey = 'cert.summary.issue'; $sumVal = @($rows.Count, $issues) }
    else               { $sumKey = 'cert.summary.ok';    $sumVal = @($rows.Count) }

    return New-PMResult -Id 'CERT' -TitleKey 'cert.title' -Status $overall `
        -SummaryKey $sumKey -SummaryValues $sumVal `
        -Columns $columns -Rows $rows -Findings $findings -Raw $raw
}

Register-PMCheck -Id 'CERT' -TitleKey 'cert.title' -Function 'Invoke-PMCheckCertificates'
