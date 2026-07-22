# PMtools check - Top resource-consuming processes.  ASCII-only; text comes from i18n.json.
# Purely diagnostic context for the CPU/memory section: always INFO, never raises a finding.

function Invoke-PMCheckTopProcess {

    $take = [int](Get-PMSetting -Path 'Lookback.TopProcessCount' -Default 10)

    $procs = @(Get-Process -ErrorAction SilentlyContinue |
               Sort-Object -Property WorkingSet64 -Descending |
               Select-Object -First $take)

    $columns = @(
        (New-PMColumn -Key 'Name' -TextKey 'proc.col.name'),
        (New-PMColumn -Key 'Pid'  -TextKey 'proc.col.pid' -Align 'right'),
        (New-PMColumn -Key 'Cpu'  -TextKey 'proc.col.cpu' -Align 'right'),
        (New-PMColumn -Key 'Mem'  -TextKey 'proc.col.mem' -Align 'right')
    )

    $rows = @()
    $raw  = @()
    foreach ($p in $procs) {
        # CPU time is unreadable for processes owned by another account when not
        # elevated; show a dash rather than a misleading zero.
        if ($null -ne $p.CPU) { $cpu = [math]::Round($p.CPU, 1) } else { $cpu = '-' }
        $memMB = [math]::Round($p.WorkingSet64 / 1MB, 1)

        $rows += New-PMRow -Values @{ Name = $p.ProcessName; Pid = $p.Id; Cpu = $cpu; Mem = $memMB }
        $raw  += [pscustomobject]@{ Name = $p.ProcessName; Pid = $p.Id; CpuSeconds = $p.CPU; MemoryMB = $memMB }
    }

    return New-PMResult -Id 'PROC' -TitleKey 'proc.title' -Status 'INFO' `
        -SummaryKey 'proc.summary' -SummaryValues @($rows.Count) `
        -Columns $columns -Rows $rows -Findings @() -Raw $raw
}

Register-PMCheck -Id 'PROC' -TitleKey 'proc.title' -Function 'Invoke-PMCheckTopProcess'
