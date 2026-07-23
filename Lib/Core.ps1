# =====================================================================
#  PMtools - Core.ps1
#  Shared contract used by every check and by the report renderer.
#
#  IMPORTANT: this file must stay ASCII-only.
#  PowerShell 5.1 reads .ps1 files as ANSI unless they carry a UTF-8 BOM,
#  so any Thai text placed here would be corrupted. All display text lives
#  in Config\i18n.json, which is read with -Encoding UTF8 instead.
# =====================================================================

$Script:PMConfig     = $null
$Script:PMText       = @{}
$Script:PMWatchlist  = @()
$Script:PMChecks     = @()
$Script:PMStatusRank = @{ 'INFO' = 0; 'OK' = 1; 'WARN' = 2; 'ERROR' = 3; 'CRIT' = 4 }

# ---------------------------------------------------------------------
# Check registration
#   Each file in Checks\ ends with a Register-PMCheck call. Declaring the Id
#   and TitleKey up front lets the orchestrator honour -Only and still build a
#   properly titled ERROR card when the check itself throws.
# ---------------------------------------------------------------------

function Register-PMCheck {
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][string]$TitleKey,
        [Parameter(Mandatory)][string]$Function
    )
    $Script:PMChecks += [pscustomobject]@{
        Id       = $Id
        TitleKey = $TitleKey
        Function = $Function
    }
}

function Get-PMRegisteredCheck { return $Script:PMChecks }

function Clear-PMRegisteredCheck { $Script:PMChecks = @() }

# ---------------------------------------------------------------------
# Configuration loading
# ---------------------------------------------------------------------

function Import-PMJson {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Configuration file not found: $Path"
    }
    # -Encoding UTF8 is what makes the Thai text in the JSON files survive.
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "Configuration file is empty: $Path"
    }
    return ($raw | ConvertFrom-Json)
}

function Initialize-PMCore {
    param([Parameter(Mandatory)][string]$ConfigDir)

    # Remembered so anything loaded later can find its own config without the
    # caller having to thread the path through - Lib\ArcGIS.ps1 stores the
    # saved connection beside settings.json.
    $Script:PMConfigDir = $ConfigDir

    $Script:PMConfig = Import-PMJson -Path (Join-Path $ConfigDir 'settings.json')

    $i18n = Import-PMJson -Path (Join-Path $ConfigDir 'i18n.json')
    $Script:PMText = @{}
    foreach ($prop in $i18n.PSObject.Properties) {
        if ($prop.Name -like '_*') { continue }   # keys starting with _ are comments
        $Script:PMText[$prop.Name] = $prop.Value
    }

    $Script:PMWatchlist = @()
    $svcPath = Join-Path $ConfigDir 'services.json'
    if (Test-Path -LiteralPath $svcPath) {
        $svc = Import-PMJson -Path $svcPath
        if ($svc.Watchlist) { $Script:PMWatchlist = @($svc.Watchlist) }
    }
}

function Get-PMConfig { return $Script:PMConfig }

$Script:PMConfigDir = $null
function Get-PMConfigDir { return $Script:PMConfigDir }

# Set by the orchestrator so checks can find sibling data - the TREND check
# reads the sample files Start-PMMonitor.ps1 leaves under Output\_Perf.
$Script:PMOutputRoot = $null
function Set-PMOutputRoot { param([string]$Path) $Script:PMOutputRoot = $Path }
function Get-PMOutputRoot { return $Script:PMOutputRoot }

function Get-PMWatchlist { return $Script:PMWatchlist }

# Safe dotted lookup into settings.json so a missing key never crashes a check.
function Get-PMSetting {
    param(
        [Parameter(Mandatory)][string]$Path,
        [object]$Default = $null
    )
    $node = $Script:PMConfig
    foreach ($part in $Path.Split('.')) {
        if ($null -eq $node) { return $Default }
        $prop = $node.PSObject.Properties[$part]
        if ($null -eq $prop) { return $Default }
        $node = $prop.Value
    }
    if ($null -eq $node) { return $Default }
    return $node
}

# ---------------------------------------------------------------------
# Bilingual text
# ---------------------------------------------------------------------

function Format-PMString {
    param([string]$Template, [object[]]$Values, [string]$Key = '')

    if ($null -eq $Values -or $Values.Count -eq 0) { return $Template }
    try { return ($Template -f $Values) }
    catch {
        # A bad placeholder must not kill the run, but it must not pass
        # unnoticed either: swallowing this silently once shipped a caption
        # reading "{0} samples over {1} minutes" straight into the report,
        # because the caller passed one value fewer than the text expected.
        Write-Warning ("PMtools: text '{0}' has {1} placeholder(s) that {2} supplied value(s) could not fill. Showing the raw template." -f `
            $Key, ([regex]::Matches($Template, '\{\d+\}') | ForEach-Object { $_.Value } | Select-Object -Unique).Count, $Values.Count)
        return $Template
    }
}

# Returns an object with .Th and .En. Unknown keys surface as [key.name]
# so a missing translation is visible in the report rather than silently blank.
function Get-PMText {
    param(
        [Parameter(Mandatory)][string]$Key,
        [object[]]$Values
    )
    $entry = $Script:PMText[$Key]
    if ($null -eq $entry) {
        return [pscustomobject]@{ Th = "[$Key]"; En = "[$Key]" }
    }
    return [pscustomobject]@{
        Th = Format-PMString -Template ([string]$entry.th) -Values $Values -Key "$Key.th"
        En = Format-PMString -Template ([string]$entry.en) -Values $Values -Key "$Key.en"
    }
}

function New-PMBiText {
    param([string]$Th = '', [string]$En = '')
    return [pscustomobject]@{ Th = $Th; En = $En }
}

# ---------------------------------------------------------------------
# Dates
#   Header dates are spelled out per language (Thai uses the Buddhist Era).
#   Table cells use yyyy-MM-dd HH:mm (Common Era) instead: one cell holds a
#   single value for both languages, and an unlabelled 2569/2026 would be
#   ambiguous to half the readers.
# ---------------------------------------------------------------------

function Format-PMDateTime {
    param([Parameter(Mandatory)][datetime]$Date)

    $months = $Script:PMText['date.months']
    $tmpl   = $Script:PMText['date.longTime']
    if ($null -eq $months -or $null -eq $tmpl) {
        return (New-PMBiText -Th $Date.ToString('yyyy-MM-dd HH:mm') -En $Date.ToString('yyyy-MM-dd HH:mm'))
    }
    $idx  = $Date.Month - 1
    $time = $Date.ToString('HH:mm')
    return [pscustomobject]@{
        Th = Format-PMString -Template ([string]$tmpl.th) -Values @($Date.Day, $months.th[$idx], ($Date.Year + 543), $time)
        En = Format-PMString -Template ([string]$tmpl.en) -Values @($Date.Day, $months.en[$idx], $Date.Year, $time)
    }
}

function Format-PMStamp {
    param([object]$Date)

    if ($null -eq $Date) { return '' }
    try { return ([datetime]$Date).ToString('yyyy-MM-dd HH:mm') }
    catch { return [string]$Date }
}

function Get-PMAgeDays {
    param([object]$Date)

    if ($null -eq $Date) { return $null }
    try { return [math]::Round(((Get-Date) - [datetime]$Date).TotalDays, 1) }
    catch { return $null }
}

# ---------------------------------------------------------------------
# Status handling
# ---------------------------------------------------------------------

function Get-PMStatusRank {
    param([string]$Status)
    if ([string]::IsNullOrWhiteSpace($Status)) { return 0 }
    $r = $Script:PMStatusRank[$Status.ToUpper()]
    if ($null -eq $r) { return 0 }
    return $r
}

# Worst-wins roll-up, used both for a check's overall status and the report total.
function Get-PMWorstStatus {
    param([string[]]$Status)

    $worst = 'INFO'
    foreach ($s in $Status) {
        if ([string]::IsNullOrWhiteSpace($s)) { continue }
        if ((Get-PMStatusRank $s) -gt (Get-PMStatusRank $worst)) { $worst = $s.ToUpper() }
    }
    return $worst
}

# Turns a measured value into OK / WARN / CRIT using only Config\settings.json.
# Direction 'Lower' means small values are bad (free space, days until expiry);
# 'Higher' means large values are bad (CPU load, days since last patch).
function Test-PMThreshold {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][double]$Value
    )
    $t = Get-PMSetting -Path "Thresholds.$Name"
    if ($null -eq $t) { return 'INFO' }

    $warn = $t.Warn
    $crit = $t.Crit
    $lower = ([string]$t.Direction -eq 'Lower')

    if ($lower) {
        if ($null -ne $crit -and $Value -le [double]$crit) { return 'CRIT' }
        if ($null -ne $warn -and $Value -le [double]$warn) { return 'WARN' }
    }
    else {
        if ($null -ne $crit -and $Value -ge [double]$crit) { return 'CRIT' }
        if ($null -ne $warn -and $Value -ge [double]$warn) { return 'WARN' }
    }
    return 'OK'
}

function Get-PMThresholdValue {
    param(
        [Parameter(Mandatory)][string]$Name,
        [ValidateSet('Warn', 'Crit')][string]$Level = 'Warn'
    )
    return (Get-PMSetting -Path "Thresholds.$Name.$Level")
}

# ---------------------------------------------------------------------
# The check contract
#   Every check returns the same shape, so the HTML renderer never needs to
#   know which check produced the data and new checks can be added without
#   touching the renderer or the orchestrator.
# ---------------------------------------------------------------------

function New-PMColumn {
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][string]$TextKey,
        [ValidateSet('left', 'right', 'center')][string]$Align = 'left',
        [switch]$Wide
    )
    $t = Get-PMText -Key $TextKey
    return [pscustomobject]@{
        Key   = $Key
        Th    = $t.Th
        En    = $t.En
        Align = $Align
        Wide  = [bool]$Wide
    }
}

function New-PMRow {
    param(
        [Parameter(Mandatory)][hashtable]$Values,
        [string]$Status = ''
    )
    $row = @{}
    foreach ($k in $Values.Keys) { $row[$k] = $Values[$k] }
    $row['_RowStatus'] = $Status
    return $row
}

# Convenience for checks that present a simple Item / Value list.
function New-PMItemColumns {
    return @(
        (New-PMColumn -Key 'Item'  -TextKey 'ui.col.item'),
        (New-PMColumn -Key 'Value' -TextKey 'ui.col.value' -Wide)
    )
}

function New-PMItemRow {
    param(
        [Parameter(Mandatory)][string]$TextKey,
        [object]$Value,
        # Supply this when the value itself is translated (Yes/No, Enabled/Disabled).
        # The renderer picks up any "<Key>En" companion field automatically.
        [string]$ValueEn,
        [string]$Status = ''
    )
    $t = Get-PMText -Key $TextKey
    return @{
        Item       = $t.Th
        ItemEn     = $t.En
        Value      = $Value
        ValueEn    = $ValueEn
        _RowStatus = $Status
    }
}

# Shorthand for the many places a value is just a translated word.
function Get-PMWord {
    param([Parameter(Mandatory)][string]$Key)
    return (Get-PMText -Key $Key)
}

function New-PMFinding {
    param(
        [Parameter(Mandatory)][ValidateSet('CRIT', 'WARN', 'INFO', 'ERROR')][string]$Severity,
        [Parameter(Mandatory)][string]$TextKey,
        [object[]]$Values
    )
    $t = Get-PMText -Key $TextKey -Values $Values
    return [pscustomobject]@{
        Severity = $Severity
        Th       = $t.Th
        En       = $t.En
    }
}

function New-PMResult {
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][string]$TitleKey,
        [ValidateSet('OK', 'WARN', 'CRIT', 'INFO', 'ERROR')][string]$Status = 'INFO',
        [string]$SummaryKey,
        [object[]]$SummaryValues,
        [object[]]$Columns,
        [object[]]$Rows,
        [object[]]$Findings,
        # Optional plot data. The check supplies the numbers and the series
        # names; the renderer owns every presentation decision (colours, sizing,
        # the SVG itself), the same way it does for tables.
        [object]$Chart,
        [object]$Raw
    )
    $title = Get-PMText -Key $TitleKey
    if ($SummaryKey) { $summary = Get-PMText -Key $SummaryKey -Values $SummaryValues }
    else             { $summary = New-PMBiText }

    return [pscustomobject]@{
        Id         = $Id
        TitleTh    = $title.Th
        TitleEn    = $title.En
        Status     = $Status
        SummaryTh  = $summary.Th
        SummaryEn  = $summary.En
        Columns    = @($Columns)
        Rows       = @($Rows)
        Findings   = @($Findings)
        Chart      = $Chart
        Raw        = $Raw
        DurationMs = 0
    }
}

# Describes a time-series plot. Both series must share one unit and one scale:
# a second y-axis would let the chart invent a correlation that is not in the
# data, which is why CPU and memory are both carried as percentages here.
function New-PMLineChart {
    param(
        [Parameter(Mandatory)][string[]]$XLabels,
        [Parameter(Mandatory)][object[]]$Series,   # @{ TitleTh; TitleEn; Values }
        [double]$YMin = 0,
        [double]$YMax = 100,
        [string]$YUnit = '%',
        [string]$CaptionTh = '',
        [string]$CaptionEn = '',
        # Off by default so the existing CPU/Memory chart's 0/25/50/75/100%
        # gridlines are untouched. On, each gridline value is snapped to the
        # nearest round number (nearest 100 once YMax reaches 100, nearest 10
        # below that, nearest whole number below 10) instead of just an even
        # position between YMin/YMax - meant for count-style charts (request
        # volumes and the like) where a label like "7400" reads far easier
        # than the exact-but-arbitrary "7376" an even split would give.
        [switch]$RoundTicks
    )
    return [pscustomobject]@{
        Type       = 'line'
        XLabels    = @($XLabels)
        Series     = @($Series)
        YMin       = $YMin
        YMax       = $YMax
        YUnit      = $YUnit
        CaptionTh  = $CaptionTh
        CaptionEn  = $CaptionEn
        RoundTicks = [bool]$RoundTicks
    }
}

# Built by the orchestrator when a check throws, so one broken check degrades
# to a visible ERROR card instead of taking the whole report down.
function New-PMErrorResult {
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][string]$TitleKey,
        [Parameter(Mandatory)][string]$Message
    )
    $title = Get-PMText -Key $TitleKey
    return New-PMResult -Id $Id -TitleKey $TitleKey -Status 'ERROR' `
        -SummaryKey 'common.checkFailed' -SummaryValues @($Message) `
        -Findings @(New-PMFinding -Severity 'ERROR' -TextKey 'common.checkFailedReco' -Values @($title.Th)) `
        -Raw ([pscustomobject]@{ Error = $Message })
}

# ---------------------------------------------------------------------
# Misc helpers
# ---------------------------------------------------------------------

function ConvertTo-PMHtmlText {
    param([object]$Text)

    if ($null -eq $Text) { return '' }
    $s = [string]$Text
    return $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function Get-PMShortText {
    param([object]$Text, [int]$MaxLength = 160)

    if ($null -eq $Text) { return '' }
    $s = ([string]$Text) -replace '\s+', ' '
    $s = $s.Trim()
    if ($s.Length -le $MaxLength) { return $s }
    return $s.Substring(0, $MaxLength) + '...'
}

function Test-PMIsAdministrator {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $pr = New-Object Security.Principal.WindowsPrincipal($id)
        return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch { return $false }
}

function ConvertTo-PMGB {
    param([object]$Bytes, [int]$Decimals = 2)

    if ($null -eq $Bytes) { return 0 }
    return [math]::Round(([double]$Bytes / 1GB), $Decimals)
}

# ---------------------------------------------------------------------
# CPU and memory sampling
#   Kept here so the PERF check and Start-PMMonitor.ps1 read the same numbers
#   from the same source - two different sources would put contradictory
#   figures in the same report.
#
#   WMI *class* names are always English, unlike performance counter paths
#   ('\Processor(_Total)\% Processor Time' does not resolve on a non-English
#   Windows), so this stays correct whatever language the server runs in.
# ---------------------------------------------------------------------

function Get-PMCpuPercent {
    # Preferred: the formatted perf class, which is what Task Manager shows.
    try {
        $v = (Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor `
                              -Filter "Name='_Total'" -ErrorAction Stop).PercentProcessorTime
        if ($null -ne $v) { return [double]$v }
    }
    catch { }

    # Fallback for servers whose performance counters are damaged or disabled.
    try {
        $v = (Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop |
              Measure-Object -Property LoadPercentage -Average).Average
        if ($null -ne $v) { return [double]$v }
    }
    catch { }

    return $null
}

function Get-PMMemoryUsage {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $totalGB = ConvertTo-PMGB ($os.TotalVisibleMemorySize * 1KB) 2
    $freeGB  = ConvertTo-PMGB ($os.FreePhysicalMemory * 1KB) 2
    $usedGB  = [math]::Round($totalGB - $freeGB, 2)

    if ($totalGB -gt 0) { $freePct = [math]::Round((($freeGB / $totalGB) * 100), 1) } else { $freePct = 0 }

    return [pscustomobject]@{
        TotalGB     = $totalGB
        UsedGB      = $usedGB
        FreeGB      = $freeGB
        FreePercent = $freePct
        UsedPercent = [math]::Round(100 - $freePct, 1)
    }
}

# Percentile over an unsorted set. Used instead of the maximum when judging
# sustained load: a one-second spike to 100% is not a capacity problem, a 95th
# percentile of 90% is.
function Get-PMPercentile {
    param(
        [Parameter(Mandatory)][double[]]$Values,
        [double]$Percentile = 95
    )
    if ($Values.Count -eq 0) { return $null }

    $sorted = @($Values | Sort-Object)
    if ($sorted.Count -eq 1) { return $sorted[0] }

    $rank = ($Percentile / 100) * ($sorted.Count - 1)
    $low  = [math]::Floor($rank)
    $high = [math]::Ceiling($rank)
    if ($low -eq $high) { return $sorted[[int]$rank] }

    # Linear interpolation between the two neighbouring samples.
    return $sorted[[int]$low] + (($rank - $low) * ($sorted[[int]$high] - $sorted[[int]$low]))
}

function Write-PMLog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Good', 'Warn', 'Bad', 'Step')][string]$Level = 'Info'
    )
    switch ($Level) {
        'Good' { Write-Host $Message -ForegroundColor Green }
        'Warn' { Write-Host $Message -ForegroundColor Yellow }
        'Bad'  { Write-Host $Message -ForegroundColor Red }
        'Step' { Write-Host $Message -ForegroundColor Cyan }
        default { Write-Host $Message }
    }
}
