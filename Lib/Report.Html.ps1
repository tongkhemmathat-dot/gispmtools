# =====================================================================
#  PMtools - Report.Html.ps1
#  Renders the check results into one self-contained HTML file.
#
#  ASCII-only. Every visible string comes from Config\i18n.json; the few
#  symbols used are emitted as HTML numeric entities.
#
#  No CDN, no external stylesheet, no remote font, no image request: the
#  finished file opens correctly on an isolated network and can be mailed
#  as a single attachment.
# =====================================================================

$Script:PMStatusGlyph = @{
    'OK'    = '&#10003;'   # check mark
    'WARN'  = '&#9650;'    # up triangle
    'CRIT'  = '&#10007;'   # ballot x
    'INFO'  = '&#9679;'    # filled circle
    'ERROR' = '&#33;'      # exclamation
}

$Script:PMSeverityRank = @{ 'CRIT' = 0; 'ERROR' = 1; 'WARN' = 2; 'INFO' = 3 }

function Get-PMReportCss {
    return @'
:root{
  --ink:#1a2332; --muted:#5b6b7f; --faint:#8496a9;
  --line:#dde3ea; --line-soft:#eef1f5; --paper:#ffffff; --bg:#f4f6f9;
  --navy:#1f3a5f;
  --ok-fg:#0e6b3d;   --ok-bg:#e6f4ec;   --ok-line:#b8dfc9;
  --warn-fg:#8a5300; --warn-bg:#fdf4e3; --warn-line:#f0d9a8;
  --crit-fg:#a91b16; --crit-bg:#fdeceb; --crit-line:#f3c2be;
  --info-fg:#40536b; --info-bg:#eef2f7; --info-line:#d4dde8;
  --err-fg:#6b3410;  --err-bg:#fbeee4;  --err-line:#e8cdb4;
}
*{box-sizing:border-box}
body{
  margin:0; padding:28px 16px 64px;
  background:var(--bg); color:var(--ink);
  font-family:"Sarabun","TH SarabunPSK","Leelawadee UI","Segoe UI",Tahoma,sans-serif;
  font-size:15px; line-height:1.65;
}
.sheet{max-width:1120px; margin:0 auto}

/* ---------- bilingual text ---------- */
.bi{display:block}
.bi .th{display:block}
.bi .en{display:block; font-size:.82em; color:var(--faint); line-height:1.35}

/* ---------- header ---------- */
.masthead{
  background:var(--navy); color:#fff;
  border-radius:10px 10px 0 0; padding:26px 30px 22px;
}
.masthead .org{font-size:15px; opacity:.9}
.masthead .org .en{color:rgba(255,255,255,.62)}
.masthead h1{margin:12px 0 2px; font-size:26px; font-weight:700; letter-spacing:.2px}
.masthead .sub{font-size:14px; color:rgba(255,255,255,.66); margin-bottom:0}
.metagrid{
  background:var(--paper); border:1px solid var(--line); border-top:none;
  border-radius:0 0 10px 10px; padding:18px 30px 20px; margin-bottom:24px;
  display:grid; grid-template-columns:repeat(auto-fit,minmax(230px,1fr)); gap:14px 28px;
}
.metagrid .k{font-size:12px; color:var(--faint); text-transform:uppercase; letter-spacing:.6px}
.metagrid .v{font-weight:600}
.metagrid .v .en{font-weight:400}

/* ---------- sections ---------- */
section{
  background:var(--paper); border:1px solid var(--line); border-radius:10px;
  padding:22px 26px 26px; margin-bottom:22px;
}
h2{
  margin:0 0 18px; font-size:19px; font-weight:700;
  padding-bottom:11px; border-bottom:2px solid var(--navy);
}
h2 .en{font-size:.68em; font-weight:400; color:var(--faint); display:block; line-height:1.3}

/* ---------- summary ---------- */
.tiles{display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:12px; margin-bottom:18px}
.tile{border:1px solid var(--line); border-radius:8px; padding:14px 16px; border-left-width:5px}
.tile .n{font-size:30px; font-weight:700; line-height:1.15}
.tile .l{font-size:13px; color:var(--muted)}
.tile .l .en{font-size:.85em; color:var(--faint)}
.tile.t-OK{border-left-color:var(--ok-line);   background:var(--ok-bg)}
.tile.t-OK .n{color:var(--ok-fg)}
.tile.t-WARN{border-left-color:var(--warn-line); background:var(--warn-bg)}
.tile.t-WARN .n{color:var(--warn-fg)}
.tile.t-CRIT{border-left-color:var(--crit-line); background:var(--crit-bg)}
.tile.t-CRIT .n{color:var(--crit-fg)}
.tile.t-INFO{border-left-color:var(--info-line); background:var(--info-bg)}
.tile.t-INFO .n{color:var(--info-fg)}
.tile.t-ERROR{border-left-color:var(--err-line); background:var(--err-bg)}
.tile.t-ERROR .n{color:var(--err-fg)}
.verdict{
  border-left:4px solid var(--navy); background:var(--info-bg);
  padding:14px 18px; border-radius:0 8px 8px 0;
}
.verdict .th{font-size:16px}
.verdict .en{font-size:13px; color:var(--muted); margin-top:5px}
.notice{
  margin-top:14px; padding:11px 16px; border-radius:8px;
  background:var(--warn-bg); border:1px solid var(--warn-line); color:var(--warn-fg); font-size:14px;
}

/* ---------- tables ---------- */
.tablewrap{overflow-x:auto; -webkit-overflow-scrolling:touch}
table{border-collapse:collapse; width:100%; font-size:14px}
th,td{padding:9px 12px; border-bottom:1px solid var(--line-soft); text-align:left; vertical-align:top}
thead th{
  background:#f7f9fb; border-bottom:2px solid var(--line);
  font-size:13px; font-weight:700; white-space:nowrap;
}
thead th .en{font-size:.84em; font-weight:400; color:var(--faint); white-space:nowrap}
td.num,th.num{text-align:right; white-space:nowrap; font-variant-numeric:tabular-nums}
td.ctr,th.ctr{text-align:center}
td.wide{min-width:280px}
tbody tr:last-child td{border-bottom:none}
tbody tr.r-WARN{background:var(--warn-bg)}
tbody tr.r-CRIT{background:var(--crit-bg)}
tbody tr.r-ERROR{background:var(--err-bg)}
.empty{color:var(--faint); font-size:14px; padding:14px 2px}

/* ---------- badges ---------- */
.badge{
  display:inline-block; white-space:nowrap; padding:2px 10px; border-radius:999px;
  font-size:12.5px; font-weight:700; border:1px solid transparent;
}
.badge .g{margin-right:5px; font-weight:400}
.badge .en{font-weight:400; opacity:.75; font-size:.9em}
.b-OK{color:var(--ok-fg); background:var(--ok-bg); border-color:var(--ok-line)}
.b-WARN{color:var(--warn-fg); background:var(--warn-bg); border-color:var(--warn-line)}
.b-CRIT{color:var(--crit-fg); background:var(--crit-bg); border-color:var(--crit-line)}
.b-INFO{color:var(--info-fg); background:var(--info-bg); border-color:var(--info-line)}
.b-ERROR{color:var(--err-fg); background:var(--err-bg); border-color:var(--err-line)}

/* ---------- overview ---------- */
.overview td a{color:var(--navy); text-decoration:none; font-weight:600}
.overview td a:hover{text-decoration:underline}

/* ---------- recommendations ---------- */
ol.reco{margin:0; padding-left:0; list-style:none; counter-reset:r}
ol.reco li{
  counter-increment:r; position:relative;
  padding:13px 16px 13px 52px; margin-bottom:10px;
  border:1px solid var(--line); border-left-width:5px; border-radius:0 8px 8px 0;
}
ol.reco li::before{
  content:counter(r); position:absolute; left:16px; top:13px;
  font-weight:700; color:var(--faint); font-size:15px;
}
ol.reco li.s-CRIT{border-left-color:var(--crit-line);  background:var(--crit-bg)}
ol.reco li.s-WARN{border-left-color:var(--warn-line);  background:var(--warn-bg)}
ol.reco li.s-ERROR{border-left-color:var(--err-line);  background:var(--err-bg)}
ol.reco li.s-INFO{border-left-color:var(--info-line);  background:var(--info-bg)}
ol.reco .head{margin-bottom:5px}
ol.reco .body .en{font-size:.86em; color:var(--muted); margin-top:4px; display:block}

/* ---------- detail cards ---------- */
.card{border:1px solid var(--line); border-radius:9px; margin-bottom:16px; overflow:hidden}
.card > .top{
  display:flex; flex-wrap:wrap; gap:10px; align-items:flex-start; justify-content:space-between;
  padding:14px 18px; background:#f7f9fb; border-bottom:1px solid var(--line);
}
.card .top h3{margin:0; font-size:16px; font-weight:700}
.card .top h3 .en{font-size:.76em; font-weight:400; color:var(--faint); display:block; line-height:1.3}
.card .note{padding:13px 18px; color:var(--muted); border-bottom:1px solid var(--line-soft); font-size:14px}
.card .note .en{font-size:.86em; color:var(--faint); display:block}
.card .body{padding:4px 6px 6px}
.anchor{scroll-margin-top:16px}

/* ---------- charts ---------- */
.chartbox{padding:16px 18px 4px}
.legend{display:flex; flex-wrap:wrap; gap:18px; margin-bottom:12px; font-size:13px}
.legend .item{display:flex; align-items:center; gap:8px}
.legend .sw{width:16px; height:3px; border-radius:2px; flex:none}
.legend .en{color:var(--faint); font-size:.86em; margin-left:4px}
.plotwrap{position:relative}
.plotwrap svg{display:block; width:100%; height:auto; overflow:visible}
.grid line{stroke:#e6eaef; stroke-width:1}
.axis text{fill:var(--faint); font-size:11px}
.serieslabel{font-size:11.5px; font-weight:700}
.crosshair{stroke:#94a3b4; stroke-width:1; opacity:0}
.cursordot{opacity:0}
.charttip{
  position:absolute; pointer-events:none; opacity:0;
  transform:translate(-50%,-115%); transition:opacity .08s;
  background:var(--paper); border:1px solid var(--line); border-radius:7px;
  padding:8px 11px; font-size:12.5px; white-space:nowrap; z-index:5;
  box-shadow:0 2px 10px rgba(16,26,40,.14);
}
.charttip .t{color:var(--faint); margin-bottom:4px}
.charttip .row{display:flex; align-items:center; gap:7px; line-height:1.5}
.charttip .sw{width:10px; height:10px; border-radius:2px; flex:none}
.charttip .v{font-weight:700; font-variant-numeric:tabular-nums; margin-left:auto; padding-left:10px}
.chartcap{font-size:12.5px; color:var(--faint); padding:10px 2px 2px}
.chartcap .en{display:block}

/* ---------- footer ---------- */
footer{color:var(--faint); font-size:12.5px; text-align:center; padding:10px 0 0}
footer .en{display:block; margin-top:3px}

/* ---------- toolbar ---------- */
.toolbar{max-width:1120px; margin:0 auto 14px; text-align:right}
.btn{
  border:1px solid var(--line); background:var(--paper); color:var(--ink);
  border-radius:7px; padding:8px 16px; font-size:14px; cursor:pointer;
  font-family:inherit;
}
.btn:hover{background:#eef2f7}

/* ---------- print ---------- */
@media print{
  @page{margin:14mm 11mm}
  .crosshair,.cursordot,.charttip{display:none !important}
  .chartbox{page-break-inside:avoid; break-inside:avoid}
  html,body{background:#fff}
  body{padding:0; font-size:11.5pt}
  *{-webkit-print-color-adjust:exact !important; print-color-adjust:exact !important}
  .toolbar{display:none !important}
  .sheet{max-width:none}
  section,.card,.metagrid,.masthead{border-radius:0}
  section{border:none; padding:0 0 8mm; margin-bottom:6mm; page-break-inside:auto}
  h2{page-break-after:avoid}
  .card{page-break-inside:avoid; break-inside:avoid}
  ol.reco li{page-break-inside:avoid; break-inside:avoid}
  tr,thead{page-break-inside:avoid}
  thead{display:table-header-group}
  .tablewrap{overflow:visible}
  a{color:inherit; text-decoration:none}
}
'@
}

# Emits a status pill carrying colour, glyph and the status word in both
# languages, so the meaning survives greyscale printing and colour blindness.
function Get-PMStatusBadge {
    param([string]$Status)

    $s = $Status.ToUpper()
    $t = Get-PMText -Key "status.$s"
    $g = $Script:PMStatusGlyph[$s]
    if (-not $g) { $g = '' }
    return ('<span class="badge b-{0}"><span class="g">{1}</span>{2} <span class="en">/ {3}</span></span>' -f `
        $s, $g, (ConvertTo-PMHtmlText $t.Th), (ConvertTo-PMHtmlText $t.En))
}

function Get-PMBiBlock {
    param([string]$Th, [string]$En, [string]$Class = 'bi')

    $out = '<span class="' + $Class + '"><span class="th">' + (ConvertTo-PMHtmlText $Th) + '</span>'
    if (-not [string]::IsNullOrWhiteSpace($En)) {
        $out += '<span class="en">' + (ConvertTo-PMHtmlText $En) + '</span>'
    }
    return $out + '</span>'
}

# Renders one result's data table. Any column key K is rendered bilingually when
# the row also carries "<K>En" - that is how Item/ItemEn and Result/ResultEn
# reach the page without the renderer knowing anything about those checks.
function Get-PMResultTable {
    param([object]$Result)

    if (-not $Result.Columns -or $Result.Columns.Count -eq 0) { return '' }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="tablewrap"><table><thead><tr>')

    foreach ($col in $Result.Columns) {
        $cls = ''
        if ($col.Align -eq 'right')  { $cls = ' class="num"' }
        if ($col.Align -eq 'center') { $cls = ' class="ctr"' }
        [void]$sb.Append('<th' + $cls + '>' + (Get-PMBiBlock -Th $col.Th -En $col.En) + '</th>')
    }
    [void]$sb.Append('<th class="ctr">' + (Get-PMBiBlock -Th (Get-PMText -Key 'ui.col.status').Th -En (Get-PMText -Key 'ui.col.status').En) + '</th>')
    [void]$sb.Append('</tr></thead><tbody>')

    foreach ($row in $Result.Rows) {
        $rowStatus = [string]$row['_RowStatus']
        if ([string]::IsNullOrWhiteSpace($rowStatus)) { $rowClass = '' } else { $rowClass = ' class="r-' + $rowStatus.ToUpper() + '"' }
        [void]$sb.Append('<tr' + $rowClass + '>')

        foreach ($col in $Result.Columns) {
            $cls = ''
            if ($col.Align -eq 'right')  { $cls = ' class="num"' }
            if ($col.Align -eq 'center') { $cls = ' class="ctr"' }
            elseif ($col.Wide)           { $cls = ' class="wide"' }

            $value   = $row[$col.Key]
            $enValue = $row[($col.Key + 'En')]

            if ($null -ne $enValue -and -not [string]::IsNullOrWhiteSpace([string]$enValue)) {
                $cell = Get-PMBiBlock -Th ([string]$value) -En ([string]$enValue)
            }
            else {
                $cell = ConvertTo-PMHtmlText $value
            }
            [void]$sb.Append('<td' + $cls + '>' + $cell + '</td>')
        }

        if ([string]::IsNullOrWhiteSpace($rowStatus)) { $badge = '' } else { $badge = Get-PMStatusBadge -Status $rowStatus }
        [void]$sb.Append('<td class="ctr">' + $badge + '</td></tr>')
    }

    [void]$sb.Append('</tbody></table></div>')
    return $sb.ToString()
}

# Categorical series colours, assigned by slot in fixed order and never cycled.
# Validated against this report's white surface with the dataviz palette
# validator: adjacent CVD deltaE 26.5, normal-vision 29.0, both above the gate,
# all slots above 3:1 contrast. The tritan separation (7.6) sits in the band
# that requires a second channel besides hue, which is why every line carries a
# direct label at its right end in addition to the legend.
$Script:PMSeriesColor = @('#2a78d6', '#008300', '#eb6834', '#4a3aa7')

# Renders a time series as inline SVG: no chart library, no remote request, and
# it prints correctly. One shared y-axis for every series - a second y-scale
# would let the chart imply a correlation that is not in the data - which is why
# the check hands over CPU and memory both expressed as percentages.
function New-PMChartHtml {
    param([object]$Result)

    $chart = $Result.Chart
    if ($null -eq $chart) { return '' }
    if ($chart.Series.Count -eq 0) { return '' }

    $n = @($chart.Series[0].Values).Count
    if ($n -lt 2) { return '' }

    # viewBox units; the SVG then scales fluidly to the card width.
    $vbW = 920; $vbH = 300
    # The right margin holds the direct end labels. It is sized for a Thai
    # series name plus its value: Thai names run considerably longer than their
    # English equivalents, and at 104 units the memory label was clipped by the
    # card edge.
    $mL = 46; $mR = 168; $mT = 14; $mB = 40
    $plotW = $vbW - $mL - $mR
    $plotH = $vbH - $mT - $mB
    $plotR = $mL + $plotW
    $plotB = $mT + $plotH

    $yMin = [double]$chart.YMin
    $yMax = [double]$chart.YMax
    $range = $yMax - $yMin
    if ($range -le 0) { $range = 1 }

    # Precomputed positions are emitted with the data so the hover layer never
    # has to re-derive the scales in JavaScript.
    $xs = @()
    for ($i = 0; $i -lt $n; $i++) {
        $xs += [math]::Round(($mL + ($i / ($n - 1)) * $plotW), 1)
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="chartbox">')

    # ---- legend (always present for two or more series) ----
    [void]$sb.Append('<div class="legend">')
    for ($s = 0; $s -lt $chart.Series.Count; $s++) {
        $color = $Script:PMSeriesColor[$s % $Script:PMSeriesColor.Count]
        [void]$sb.Append('<div class="item"><span class="sw" style="background:' + $color + '"></span>' +
            '<span>' + (ConvertTo-PMHtmlText $chart.Series[$s].TitleTh) + '</span>' +
            '<span class="en">' + (ConvertTo-PMHtmlText $chart.Series[$s].TitleEn) + '</span></div>')
    }
    [void]$sb.Append('</div>')

    [void]$sb.Append('<div class="plotwrap">')
    [void]$sb.Append('<svg viewBox="0 0 ' + $vbW + ' ' + $vbH + '" role="img" preserveAspectRatio="xMidYMid meet">')

    # ---- horizontal grid and y ticks (solid hairlines, never dashed) ----
    # Five values evenly spaced across YMin..YMax, not a hardcoded 0/25/50/75/100 -
    # that literal set only ever made sense for a 0-100 percent scale. For
    # YMin=0/YMax=100 (every chart before this one) this reproduces the exact
    # same five values, so existing charts render identically.
    # Rounded to whole numbers - a decimal on an axis label (e.g. "7375.8")
    # reads as false precision the underlying data does not actually have.
    [void]$sb.Append('<g class="grid">')
    $ticks = @()
    if ($chart.PSObject.Properties['RoundTicks'] -and $chart.RoundTicks) {
        # Snap each gridline to a round number rather than an even position -
        # opt-in only (see New-PMLineChart), so the CPU/Memory chart's
        # 0/25/50/75/100% is untouched. The round-to unit itself scales with
        # the chart's own magnitude: a request-count chart in the tens of
        # thousands reads better rounded to the nearest 100 than the nearest 1.
        if     ($yMax -ge 100) { $roundTo = 100 }
        elseif ($yMax -ge 10)  { $roundTo = 10 }
        else                   { $roundTo = 1 }
        for ($i = 0; $i -le 4; $i++) {
            $raw = $yMin + ($range * $i / 4)
            $ticks += [math]::Round($raw / $roundTo) * $roundTo
        }
    }
    else {
        for ($i = 0; $i -le 4; $i++) { $ticks += [math]::Round($yMin + ($range * $i / 4), 0) }
    }
    foreach ($t in $ticks) {
        $y = [math]::Round(($mT + (1 - (($t - $yMin) / $range)) * $plotH), 1)
        [void]$sb.Append('<line x1="' + $mL + '" y1="' + $y + '" x2="' + $plotR + '" y2="' + $y + '"/>')
    }
    [void]$sb.Append('</g><g class="axis">')
    foreach ($t in $ticks) {
        $y = [math]::Round(($mT + (1 - (($t - $yMin) / $range)) * $plotH), 1)
        [void]$sb.Append('<text x="' + ($mL - 9) + '" y="' + ($y + 4) + '" text-anchor="end">' + $t + $chart.YUnit + '</text>')
    }

    # ---- x labels: about six, with the outer two anchored inward so they
    #      cannot overflow the plot ----
    $labelCount = [math]::Min(6, $n)
    for ($k = 0; $k -lt $labelCount; $k++) {
        if ($labelCount -eq 1) { $idx = 0 } else { $idx = [int][math]::Round((($n - 1) * $k) / ($labelCount - 1)) }
        if ($k -eq 0)                    { $anchor = 'start' }
        elseif ($k -eq $labelCount - 1)  { $anchor = 'end' }
        else                             { $anchor = 'middle' }
        [void]$sb.Append('<text x="' + $xs[$idx] + '" y="' + ($plotB + 20) + '" text-anchor="' + $anchor + '">' +
            (ConvertTo-PMHtmlText $chart.XLabels[$idx]) + '</text>')
    }
    [void]$sb.Append('</g>')

    # ---- the series themselves ----
    $seriesJson = @()
    $endLabels  = @()

    for ($s = 0; $s -lt $chart.Series.Count; $s++) {

        $color  = $Script:PMSeriesColor[$s % $Script:PMSeriesColor.Count]
        $values = @($chart.Series[$s].Values)
        $ys     = @()
        $pts    = @()

        for ($i = 0; $i -lt $n; $i++) {
            $v = [double]$values[$i]
            $y = [math]::Round(($mT + (1 - (($v - $yMin) / $range)) * $plotH), 1)
            $ys  += $y
            $pts += ("{0},{1}" -f $xs[$i], $y)
        }

        [void]$sb.Append('<polyline fill="none" stroke="' + $color + '" stroke-width="2" ' +
            'stroke-linejoin="round" stroke-linecap="round" points="' + ($pts -join ' ') + '"/>')

        $endLabels += [pscustomobject]@{
            Y     = $ys[$n - 1]
            Color = $color
            Text  = ("{0} {1}{2}" -f $chart.Series[$s].TitleTh, [math]::Round([double]$values[$n - 1], 1), $chart.YUnit)
        }

        $seriesJson += [pscustomobject]@{
            name   = $chart.Series[$s].TitleTh
            nameEn = $chart.Series[$s].TitleEn
            color  = $color
            values = @($values | ForEach-Object { [math]::Round([double]$_, 1) })
            ys     = $ys
        }
    }

    # ---- direct end labels: the second channel besides hue, and the reason a
    #      reader never has to consult the legend to identify a line ----
    $sorted = @($endLabels | Sort-Object Y)
    for ($i = 1; $i -lt $sorted.Count; $i++) {
        if (($sorted[$i].Y - $sorted[$i - 1].Y) -lt 14) { $sorted[$i].Y = $sorted[$i - 1].Y + 14 }
    }
    foreach ($lab in $sorted) {
        [void]$sb.Append('<text class="serieslabel" x="' + ($plotR + 9) + '" y="' + ($lab.Y + 4) + '" fill="' + $lab.Color + '">' +
            (ConvertTo-PMHtmlText $lab.Text) + '</text>')
    }

    # ---- hover layer: one full-height capture rect, so the target is the whole
    #      column rather than a pinpoint on the line ----
    [void]$sb.Append('<line class="crosshair" x1="0" y1="' + $mT + '" x2="0" y2="' + $plotB + '"/>')
    for ($s = 0; $s -lt $chart.Series.Count; $s++) {
        $color = $Script:PMSeriesColor[$s % $Script:PMSeriesColor.Count]
        [void]$sb.Append('<circle class="cursordot" data-s="' + $s + '" r="4.5" fill="' + $color + '" stroke="#ffffff" stroke-width="2"/>')
    }
    [void]$sb.Append('<rect class="hit" x="' + $mL + '" y="' + $mT + '" width="' + $plotW + '" height="' + $plotH + '" fill="transparent"/>')
    [void]$sb.Append('</svg>')

    [void]$sb.Append('<div class="charttip"></div>')

    $payload = [pscustomobject]@{
        vbW = $vbW; vbH = $vbH
        xs = $xs
        labels = @($chart.XLabels)
        unit = $chart.YUnit
        series = $seriesJson
    }
    $json = $payload | ConvertTo-Json -Depth 6 -Compress
    # Only the closing-tag sequence can break out of a script element.
    $json = $json.Replace('</', '<\/')
    [void]$sb.Append('<script type="application/json" class="chartdata">' + $json + '</script>')

    [void]$sb.Append('</div>')   # .plotwrap

    if (-not [string]::IsNullOrWhiteSpace($chart.CaptionTh)) {
        [void]$sb.Append('<div class="chartcap">' + (ConvertTo-PMHtmlText $chart.CaptionTh) +
            '<span class="en">' + (ConvertTo-PMHtmlText $chart.CaptionEn) + '</span></div>')
    }

    [void]$sb.Append('</div>')   # .chartbox
    return $sb.ToString()
}

function Get-PMChartScript {
    return @'
(function () {
  document.querySelectorAll(".plotwrap").forEach(function (wrap) {
    var node = wrap.querySelector("script.chartdata");
    if (!node) { return; }
    var d = JSON.parse(node.textContent);

    var svg  = wrap.querySelector("svg");
    var hit  = wrap.querySelector("rect.hit");
    var line = wrap.querySelector(".crosshair");
    var dots = wrap.querySelectorAll(".cursordot");
    var tip  = wrap.querySelector(".charttip");

    function nearest(clientX) {
      var box = svg.getBoundingClientRect();
      var x = ((clientX - box.left) / box.width) * d.vbW;
      var best = 0, bestDist = Infinity;
      for (var i = 0; i < d.xs.length; i++) {
        var dist = Math.abs(d.xs[i] - x);
        if (dist < bestDist) { bestDist = dist; best = i; }
      }
      return best;
    }

    function show(i) {
      line.setAttribute("x1", d.xs[i]);
      line.setAttribute("x2", d.xs[i]);
      line.style.opacity = 1;

      var top = d.vbH;
      dots.forEach(function (dot) {
        var s = d.series[parseInt(dot.getAttribute("data-s"), 10)];
        dot.setAttribute("cx", d.xs[i]);
        dot.setAttribute("cy", s.ys[i]);
        dot.style.opacity = 1;
        if (s.ys[i] < top) { top = s.ys[i]; }
      });

      var html = '<div class="t">' + d.labels[i] + "</div>";
      d.series.forEach(function (s) {
        html += '<div class="row"><span class="sw" style="background:' + s.color + '"></span>' +
                "<span>" + s.name + '</span><span class="v">' + s.values[i] + d.unit + "</span></div>";
      });
      tip.innerHTML = html;
      tip.style.left = (d.xs[i] / d.vbW * 100) + "%";
      tip.style.top  = (top / d.vbH * 100) + "%";
      tip.style.opacity = 1;
    }

    function hide() {
      line.style.opacity = 0;
      dots.forEach(function (dot) { dot.style.opacity = 0; });
      tip.style.opacity = 0;
    }

    hit.addEventListener("mousemove", function (e) { show(nearest(e.clientX)); });
    hit.addEventListener("mouseleave", hide);
    hit.addEventListener("touchmove", function (e) {
      if (e.touches.length) { show(nearest(e.touches[0].clientX)); }
    }, { passive: true });
    hit.addEventListener("touchend", hide);
  });
})();
'@
}

function New-PMHtmlReport {
    param(
        [Parameter(Mandatory)][object[]]$Results,
        [Parameter(Mandatory)][object]$Meta
    )

    $cfg  = $Meta.Config
    $date = Format-PMDateTime -Date $Meta.GeneratedAt

    $counts = @{ OK = 0; WARN = 0; CRIT = 0; INFO = 0; ERROR = 0 }
    foreach ($r in $Results) { $counts[$r.Status] = $counts[$r.Status] + 1 }

    $sb = New-Object System.Text.StringBuilder

    # ---------- document head ----------
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="th">')
    [void]$sb.AppendLine('<head>')
    [void]$sb.AppendLine('<meta charset="utf-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
    [void]$sb.AppendLine('<title>' + (ConvertTo-PMHtmlText ((Get-PMText -Key 'ui.reportTitle').Th + ' - ' + $Meta.Hostname)) + '</title>')
    [void]$sb.AppendLine('<style>' + (Get-PMReportCss) + '</style>')
    [void]$sb.AppendLine('</head><body>')

    [void]$sb.AppendLine('<div class="toolbar"><button class="btn" onclick="window.print()">' +
        (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.printButton').Th) + ' / ' +
        (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.printButton').En) + '</button></div>')

    [void]$sb.AppendLine('<div class="sheet">')

    # ---------- masthead ----------
    [void]$sb.AppendLine('<div class="masthead">')
    [void]$sb.AppendLine('<div class="org">' + (Get-PMBiBlock -Th ([string]$cfg.Organization.NameTh) -En ([string]$cfg.Organization.NameEn)) + '</div>')
    [void]$sb.AppendLine('<h1>' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.reportTitle').Th) + '</h1>')
    [void]$sb.AppendLine('<div class="sub">' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.reportTitle').En) + ' &mdash; ' +
        (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.reportSubtitle').En) + '</div>')
    [void]$sb.AppendLine('</div>')

    # ---------- meta grid ----------
    [void]$sb.AppendLine('<div class="metagrid">')
    $metaRows = @()
    $metaRows += @{ K = 'ui.hostname'; Th = $Meta.Hostname; En = '' }
    if (-not [string]::IsNullOrWhiteSpace([string]$cfg.Report.SystemNameTh)) {
        $metaRows += @{ K = 'ui.systemName'; Th = [string]$cfg.Report.SystemNameTh; En = [string]$cfg.Report.SystemNameEn }
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$cfg.Organization.DepartmentTh)) {
        $metaRows += @{ K = 'ui.department'; Th = [string]$cfg.Organization.DepartmentTh; En = [string]$cfg.Organization.DepartmentEn }
    }
    $metaRows += @{ K = 'ui.checkDate'; Th = $date.Th; En = $date.En }
    if (-not [string]::IsNullOrWhiteSpace([string]$cfg.Report.PreparedBy)) {
        $metaRows += @{ K = 'ui.preparedBy'; Th = [string]$cfg.Report.PreparedBy; En = '' }
    }
    foreach ($m in $metaRows) {
        $label = Get-PMText -Key $m.K
        [void]$sb.AppendLine('<div><div class="k">' + (ConvertTo-PMHtmlText $label.Th) + ' / ' + (ConvertTo-PMHtmlText $label.En) + '</div>' +
            '<div class="v">' + (Get-PMBiBlock -Th $m.Th -En $m.En) + '</div></div>')
    }
    [void]$sb.AppendLine('</div>')

    # ---------- executive summary ----------
    [void]$sb.AppendLine('<section><h2>' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.summaryHeading').Th) +
        '<span class="en">' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.summaryHeading').En) + '</span></h2>')

    [void]$sb.AppendLine('<div class="tiles">')
    foreach ($s in @('CRIT', 'WARN', 'OK', 'INFO', 'ERROR')) {
        if ($counts[$s] -eq 0 -and ($s -eq 'ERROR' -or $s -eq 'INFO')) { continue }
        $lab = Get-PMText -Key "status.$s"
        [void]$sb.AppendLine('<div class="tile t-' + $s + '"><div class="n">' + $counts[$s] + '</div>' +
            '<div class="l">' + (Get-PMBiBlock -Th $lab.Th -En $lab.En) + '</div></div>')
    }
    [void]$sb.AppendLine('</div>')

    $total = $Results.Count
    if ($counts.CRIT -gt 0 -or $counts.ERROR -gt 0) {
        $verdict = Get-PMText -Key 'ui.overall.crit' -Values @(($counts.CRIT + $counts.ERROR), $counts.WARN, $total)
    }
    elseif ($counts.WARN -gt 0) {
        $verdict = Get-PMText -Key 'ui.overall.warn' -Values @($counts.WARN, $total)
    }
    else {
        $verdict = Get-PMText -Key 'ui.overall.ok' -Values @($total)
    }
    [void]$sb.AppendLine('<div class="verdict"><div class="th">' + (ConvertTo-PMHtmlText $verdict.Th) + '</div>' +
        '<div class="en">' + (ConvertTo-PMHtmlText $verdict.En) + '</div></div>')

    if (-not $Meta.IsAdmin) {
        $w = Get-PMText -Key 'ui.notAdminWarning'
        [void]$sb.AppendLine('<div class="notice">' + (ConvertTo-PMHtmlText $w.Th) + ' / ' + (ConvertTo-PMHtmlText $w.En) + '</div>')
    }
    [void]$sb.AppendLine('</section>')

    # ---------- overview ----------
    [void]$sb.AppendLine('<section><h2>' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.overviewHeading').Th) +
        '<span class="en">' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.overviewHeading').En) + '</span></h2>')
    [void]$sb.AppendLine('<div class="tablewrap"><table class="overview"><thead><tr>')
    foreach ($k in @('ui.col.topic', 'ui.col.status', 'ui.col.summary')) {
        $c = Get-PMText -Key $k
        $cls = ''
        if ($k -eq 'ui.col.status') { $cls = ' class="ctr"' }
        [void]$sb.Append('<th' + $cls + '>' + (Get-PMBiBlock -Th $c.Th -En $c.En) + '</th>')
    }
    [void]$sb.AppendLine('</tr></thead><tbody>')
    foreach ($r in $Results) {
        [void]$sb.AppendLine('<tr><td><a href="#chk-' + $r.Id + '">' + (Get-PMBiBlock -Th $r.TitleTh -En $r.TitleEn) + '</a></td>' +
            '<td class="ctr">' + (Get-PMStatusBadge -Status $r.Status) + '</td>' +
            '<td>' + (Get-PMBiBlock -Th $r.SummaryTh -En $r.SummaryEn) + '</td></tr>')
    }
    [void]$sb.AppendLine('</tbody></table></div></section>')

    # ---------- recommendations ----------
    # Off by default (Report.ShowRecommendations): the full bilingual wording of
    # every finding made the document far longer than the reader wanted. The
    # findings themselves are still computed and always written to PM-Data.json,
    # so nothing is lost by hiding the section - and the per-check summary line
    # still states what was wrong in one sentence.
    if (Get-PMSetting -Path 'Report.ShowRecommendations' -Default $false) {

        $findings = @()
        foreach ($r in $Results) {
            foreach ($f in $r.Findings) {
                if ($null -eq $f) { continue }
                $findings += [pscustomobject]@{
                    Severity = $f.Severity; Th = $f.Th; En = $f.En
                    TitleTh = $r.TitleTh; TitleEn = $r.TitleEn
                    Rank = $Script:PMSeverityRank[$f.Severity]
                }
            }
        }
        $findings = @($findings | Sort-Object Rank, TitleTh)

        [void]$sb.AppendLine('<section><h2>' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.recoHeading').Th) +
            '<span class="en">' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.recoHeading').En) + '</span></h2>')
        if ($findings.Count -eq 0) {
            $n = Get-PMText -Key 'ui.recoNone'
            [void]$sb.AppendLine('<div class="verdict"><div class="th">' + (ConvertTo-PMHtmlText $n.Th) + '</div>' +
                '<div class="en">' + (ConvertTo-PMHtmlText $n.En) + '</div></div>')
        }
        else {
            [void]$sb.AppendLine('<ol class="reco">')
            foreach ($f in $findings) {
                [void]$sb.AppendLine('<li class="s-' + $f.Severity + '"><div class="head">' + (Get-PMStatusBadge -Status $f.Severity) +
                    ' <strong>' + (ConvertTo-PMHtmlText $f.TitleTh) + '</strong></div>' +
                    '<div class="body">' + (ConvertTo-PMHtmlText $f.Th) +
                    '<span class="en">' + (ConvertTo-PMHtmlText $f.En) + '</span></div></li>')
            }
            [void]$sb.AppendLine('</ol>')
        }
        [void]$sb.AppendLine('</section>')
    }

    # ---------- details ----------
    [void]$sb.AppendLine('<section><h2>' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.detailHeading').Th) +
        '<span class="en">' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.detailHeading').En) + '</span></h2>')

    foreach ($r in $Results) {
        [void]$sb.AppendLine('<div class="card anchor" id="chk-' + $r.Id + '">')
        [void]$sb.AppendLine('<div class="top"><h3>' + (ConvertTo-PMHtmlText $r.TitleTh) +
            '<span class="en">' + (ConvertTo-PMHtmlText $r.TitleEn) + '</span></h3>' +
            (Get-PMStatusBadge -Status $r.Status) + '</div>')

        if (-not [string]::IsNullOrWhiteSpace($r.SummaryTh)) {
            [void]$sb.AppendLine('<div class="note">' + (ConvertTo-PMHtmlText $r.SummaryTh) +
                '<span class="en">' + (ConvertTo-PMHtmlText $r.SummaryEn) + '</span></div>')
        }

        if ($r.Chart) { [void]$sb.Append((New-PMChartHtml -Result $r)) }

        [void]$sb.Append('<div class="body">')
        if ($r.Rows -and $r.Rows.Count -gt 0) {
            [void]$sb.Append((Get-PMResultTable -Result $r))
        }
        else {
            [void]$sb.Append('<div class="empty">' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.noRows').Th) +
                ' / ' + (ConvertTo-PMHtmlText (Get-PMText -Key 'ui.noRows').En) + '</div>')
        }
        [void]$sb.AppendLine('</div></div>')
    }
    [void]$sb.AppendLine('</section>')

    # ---------- footer ----------
    $foot = Get-PMText -Key 'ui.generatedFooter' -Values @($Meta.ToolVersion, $date.Th, $Meta.DurationSec)
    $footEn = Get-PMText -Key 'ui.generatedFooter' -Values @($Meta.ToolVersion, $date.En, $Meta.DurationSec)
    [void]$sb.AppendLine('<footer>' + (ConvertTo-PMHtmlText $foot.Th) +
        '<span class="en">' + (ConvertTo-PMHtmlText $footEn.En) + '</span></footer>')

    [void]$sb.AppendLine('</div>')

    # Only emitted when the report actually contains a chart, so a run without
    # sampling data stays a pure static document.
    $hasChart = @($Results | Where-Object { $_.Chart }).Count -gt 0
    if ($hasChart) { [void]$sb.AppendLine('<script>' + (Get-PMChartScript) + '</script>') }

    [void]$sb.AppendLine('</body></html>')
    return $sb.ToString()
}
