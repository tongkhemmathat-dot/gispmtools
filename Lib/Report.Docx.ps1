# =====================================================================
#  PMtools - Report.Docx.ps1
#  Renders the check results into a Word document via COM automation.
#
#  ASCII-only, same rule as every other .ps1 here - see Core.ps1. Every
#  visible string comes from Config\i18n.json through Get-PMText, exactly
#  as Report.Html.ps1 does; this file adds no new wording of its own.
#
#  Unlike Report.Html.ps1, which returns one big string, this file mutates a
#  live Word document through COM. Requires Microsoft Word on the machine
#  that runs it - meant for an administrator's own workstation, not the
#  server being assessed. See Export-PMDocxReport.ps1 for the entry point.
#
#  Thai is a Complex Script language in Word's font model (like Arabic and
#  Hebrew, despite not being right-to-left), so a font assigned only through
#  Font.Name/NameAscii is not enough - plain Latin runs fall back to the
#  style's theme font unless NameOther and NameBi are set too. NameFarEast
#  is skipped: it throws when the East Asian language pack is not installed,
#  and Thai does not need it. All three findings here came from driving Word
#  through this file's own test runs, not from documentation.
# =====================================================================

# COM's RGB() macro: a Windows colour is packed R + G*256 + B*65536 (BGR
# order), the reverse of how a hex triplet like #1f3a5f reads left to right.
function Get-PMWordColor {
    param([Parameter(Mandatory)][int]$R, [Parameter(Mandatory)][int]$G, [Parameter(Mandatory)][int]$B)
    return $R + ($G * 256) + ($B * 65536)
}

# Same palette as Report.Html.ps1's :root custom properties, converted by hand
# since Word COM has no CSS to read them from.
function Get-PMDocxStatusColor {
    param([Parameter(Mandatory)][string]$Status)
    switch ($Status.ToUpper()) {
        'OK'    { return @{ FG = (Get-PMWordColor 14 107 61);  BG = (Get-PMWordColor 230 244 236) } }
        'WARN'  { return @{ FG = (Get-PMWordColor 138 83 0);   BG = (Get-PMWordColor 253 244 227) } }
        'CRIT'  { return @{ FG = (Get-PMWordColor 169 27 22);  BG = (Get-PMWordColor 253 236 235) } }
        'INFO'  { return @{ FG = (Get-PMWordColor 64 83 107);  BG = (Get-PMWordColor 238 242 247) } }
        'ERROR' { return @{ FG = (Get-PMWordColor 107 52 16);  BG = (Get-PMWordColor 251 238 228) } }
        default { return @{ FG = (Get-PMWordColor 64 83 107);  BG = (Get-PMWordColor 238 242 247) } }
    }
}

$Script:PMDocxNavy  = Get-PMWordColor 31 58 95
$Script:PMDocxMuted = Get-PMWordColor 91 107 127
$Script:PMDocxFaint = Get-PMWordColor 132 150 169
$Script:PMDocxLine  = Get-PMWordColor 221 227 234
$Script:PMDocxHead  = Get-PMWordColor 247 249 251

# wdBuiltinStyle indices (Normal=-1, Heading1=-2, Heading2=-3, Title=-63).
# Looked up by this negative index rather than by name, because Styles.Item()
# by name fails on a non-English Word UI where "Heading 1" is not the style's
# actual name - the negative index is the one locale-independent handle Word
# exposes for its built-in styles.
function Set-PMWordStyleFont {
    param([Parameter(Mandatory)]$Style, [Parameter(Mandatory)][string]$Name, [int]$Size)
    $Style.Font.NameAscii = $Name
    try { $Style.Font.NameFarEast = $Name } catch { }
    $Style.Font.NameOther = $Name
    $Style.Font.NameBi    = $Name
    if ($Size) { $Style.Font.Size = $Size }
}

# Appends one paragraph at the end of the document. Doc.Content is re-fetched
# and re-collapsed on every call rather than carrying a cursor Range between
# calls, because a Range captured before a table insert does not reliably
# follow the document past it - re-deriving "the end" from Content each time
# is what stayed correct through every shape of content this file inserts.
#
# -Tight is for a line that continues the one just written (an English
# translation under its Thai line, the way .bi .en sits close under .th in
# the HTML report) rather than starting a new block. The paragraph break
# itself is never skipped - an earlier version skipped it to save the extra
# gap and instead glued two languages into one run with no separator at all,
# visible only once real bilingual text was rendered and the join fell mid-
# word. Word's Normal style carries 8pt space-after by default, which reads
# as a paragraph gap even between two single-line paragraphs; -Tight zeroes
# that spacing on the new paragraph instead of removing the paragraph mark.
function Add-PMWordParagraph {
    param(
        [Parameter(Mandatory)]$Doc,
        [string]$Text = '',
        [Nullable[int]]$Style,
        [Nullable[int]]$Size,
        [Nullable[int]]$Color,
        [switch]$Bold,
        [switch]$Tight
    )
    $r = $Doc.Content; $r.Collapse(0) | Out-Null
    $r.InsertParagraphAfter()
    $r = $Doc.Content; $r.Collapse(0) | Out-Null
    if ($Tight) { $r.ParagraphFormat.SpaceBefore = 0; $r.ParagraphFormat.SpaceAfter = 0 }
    if ($null -ne $Style) { $r.Style = $Style }
    if ($null -ne $Size)  { $r.Font.Size = $Size }
    if ($null -ne $Color) { $r.Font.Color = $Color }
    if ($Bold)             { $r.Font.Bold = -1 }
    $r.Text = $Text
    return $r
}

# Fills one table cell with one paragraph per entry in $Lines (each a hash of
# Text/Size/Bold/Color; only Text is required). A cell's Range.End sits past
# the end-of-cell marker, so every line after the first must MoveEnd(-1)
# before collapsing and inserting further text - skipping that silently
# corrupts the table (an extra row, or text bleeding into the next cell)
# rather than throwing, which is why every multi-line cell in this file goes
# through this one checked helper instead of repeating the dance inline.
#
# The first version of the masthead hand-rolled this same sequence directly
# against a sub-range of the cell instead of the cell's own Range, and a
# stray trailing character from the first line ended up glued onto the last
# line - found by rendering an actual document, not by inspecting the code.
function Add-PMWordCellLines {
    param([Parameter(Mandatory)]$Cell, [Parameter(Mandatory)][object[]]$Lines)

    $first = $Lines[0]
    $Cell.Range.Text = [string]$first.Text
    if ($first.Size)  { $Cell.Range.Font.Size = $first.Size }
    if ($first.Bold)  { $Cell.Range.Font.Bold = -1 }
    if ($first.Color) { $Cell.Range.Font.Color = $first.Color }

    for ($i = 1; $i -lt $Lines.Count; $i++) {
        $ln = $Lines[$i]
        $r = $Cell.Range
        $r.MoveEnd(1, -1) | Out-Null   # wdCharacter = 1; step off the end-of-cell marker
        $r.Collapse(0) | Out-Null      # wdCollapseEnd = 0
        $r.InsertParagraphAfter()
        $r.Collapse(0) | Out-Null
        $r2 = $r.Duplicate
        if ($ln.Size)  { $r2.Font.Size = $ln.Size }
        if ($ln.Bold)  { $r2.Font.Bold = -1 } else { $r2.Font.Bold = 0 }
        if ($ln.Color) { $r2.Font.Color = $ln.Color }
        $r2.InsertAfter([string]$ln.Text)
    }
}

# Thai on the first line, English smaller and grey on a second - the same
# visual relationship Report.Html.ps1 gives .bi/.en. A thin wrapper over
# Add-PMWordCellLines for the common two-line case used throughout the
# result tables.
function Set-PMWordCellBilingual {
    param(
        [Parameter(Mandatory)]$Cell,
        [string]$Th = '',
        [string]$En = '',
        [switch]$Bold
    )
    $lines = @(@{ Text = $Th; Bold = [bool]$Bold })
    if (-not [string]::IsNullOrWhiteSpace($En)) {
        $lines += @{ Text = $En; Size = 9; Color = $Script:PMDocxFaint }
    }
    Add-PMWordCellLines -Cell $Cell -Lines $lines
}

# Builds one check's data table (mirrors Get-PMResultTable in Report.Html.
# ps1). $Result.Rows came through PM-Data.json, so rows are PSCustomObject,
# not the hashtable New-PMRow builds in memory - PSCustomObject's bracket
# indexer silently returns nothing for a real property (confirmed by testing
# it directly; it does not throw, which would have been easier to notice),
# so cell values are read with the dynamic dot form ($row.($col.Key)) instead.
function New-PMDocxResultTable {
    param([Parameter(Mandatory)]$Doc, [Parameter(Mandatory)]$Result)

    if (-not $Result.Columns -or $Result.Columns.Count -eq 0) { return }

    $cols = @($Result.Columns)
    $rows = @($Result.Rows)
    $statusCol = Get-PMText -Key 'ui.col.status'

    $r = $Doc.Content; $r.Collapse(0) | Out-Null
    $r.InsertParagraphAfter()
    $r = $Doc.Content; $r.Collapse(0) | Out-Null

    $nCols = $cols.Count + 1
    $nRows = $rows.Count + 1
    $tbl = $Doc.Tables.Add($r, $nRows, $nCols)
    $tbl.Borders.Enable = -1
    $tbl.Rows.AllowBreakAcrossPages = 0   # keep each row whole across a page break
    $tbl.Range.Font.Size = 9.5

    for ($c = 0; $c -lt $cols.Count; $c++) {
        Set-PMWordCellBilingual -Cell $tbl.Cell(1, $c + 1) -Th $cols[$c].Th -En $cols[$c].En -Bold
    }
    Set-PMWordCellBilingual -Cell $tbl.Cell(1, $nCols) -Th $statusCol.Th -En $statusCol.En -Bold
    $tbl.Rows.Item(1).Shading.BackgroundPatternColor = $Script:PMDocxHead
    $tbl.Rows.Item(1).HeadingFormat = -1   # repeats the header row when a table spans a page break

    for ($i = 0; $i -lt $rows.Count; $i++) {
        $row = $rows[$i]
        $rowStatus = [string]$row._RowStatus
        $wRow = $i + 2

        for ($c = 0; $c -lt $cols.Count; $c++) {
            $key   = $cols[$c].Key
            $value = $row.($key)
            $enVal = $row.($key + 'En')
            if (-not [string]::IsNullOrWhiteSpace([string]$enVal)) {
                Set-PMWordCellBilingual -Cell $tbl.Cell($wRow, $c + 1) -Th ([string]$value) -En ([string]$enVal)
            }
            else {
                $tbl.Cell($wRow, $c + 1).Range.Text = [string]$value
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($rowStatus)) {
            $t = Get-PMText -Key "status.$($rowStatus.ToUpper())"
            $sc = Get-PMDocxStatusColor -Status $rowStatus
            $badgeCell = $tbl.Cell($wRow, $nCols)
            Set-PMWordCellBilingual -Cell $badgeCell -Th $t.Th -En $t.En
            $badgeCell.Range.Font.Color = $sc.FG
            if ($rowStatus.ToUpper() -in @('WARN', 'CRIT', 'ERROR')) {
                $tbl.Rows.Item($wRow).Shading.BackgroundPatternColor = $sc.BG
            }
        }
    }
}

function New-PMDocxReport {
    param(
        [Parameter(Mandatory)][object[]]$Results,
        [Parameter(Mandatory)][object]$Meta,
        [Parameter(Mandatory)][string]$Path
    )

    $cfg  = $Meta.Config
    $date = Format-PMDateTime -Date $Meta.GeneratedAt

    $counts = @{ OK = 0; WARN = 0; CRIT = 0; INFO = 0; ERROR = 0 }
    foreach ($res in $Results) { $counts[$res.Status] = $counts[$res.Status] + 1 }

    $word = $null
    $doc  = $null
    try {
        try { $word = New-Object -ComObject Word.Application }
        catch { throw "Microsoft Word could not be started. This tool needs Word installed on the machine it runs on: $($_.Exception.Message)" }

        $word.Visible = $false
        $doc = $word.Documents.Add()

        $doc.PageSetup.TopMargin    = 42
        $doc.PageSetup.BottomMargin = 42
        $doc.PageSetup.LeftMargin   = 50
        $doc.PageSetup.RightMargin  = 50

        Set-PMWordStyleFont -Style $doc.Styles.Item(-1)  -Name 'Leelawadee UI' -Size 11
        Set-PMWordStyleFont -Style $doc.Styles.Item(-2)  -Name 'Leelawadee UI' -Size 15
        Set-PMWordStyleFont -Style $doc.Styles.Item(-3)  -Name 'Leelawadee UI' -Size 12.5
        Set-PMWordStyleFont -Style $doc.Styles.Item(-63) -Name 'Leelawadee UI' -Size 20
        $doc.Styles.Item(-2).Font.Color = $Script:PMDocxNavy
        $doc.Styles.Item(-3).Font.Color = $Script:PMDocxNavy

        # ---------- masthead ----------
        $orgTh = [string]$cfg.Organization.NameTh
        $orgEn = [string]$cfg.Organization.NameEn
        $title = Get-PMText -Key 'ui.reportTitle'
        $sub   = Get-PMText -Key 'ui.reportSubtitle'

        $r = $Doc.Content; $r.Collapse(0) | Out-Null
        $mast = $doc.Tables.Add($r, 1, 1)
        $mast.Borders.Enable = 0
        $mcell = $mast.Cell(1, 1)
        $mcell.Shading.BackgroundPatternColor = $Script:PMDocxNavy
        $mcell.TopPadding = 16; $mcell.BottomPadding = 16
        $mcell.LeftPadding = 18; $mcell.RightPadding = 18
        $white = Get-PMWordColor 255 255 255

        $mastLines = New-Object System.Collections.Generic.List[object]
        if (-not [string]::IsNullOrWhiteSpace($orgTh)) {
            $orgLine = $orgTh
            if (-not [string]::IsNullOrWhiteSpace($orgEn)) { $orgLine += ' / ' + $orgEn }
            $mastLines.Add(@{ Text = $orgLine; Size = 11; Color = $white })
        }
        $mastLines.Add(@{ Text = $title.Th; Size = 20; Bold = $true; Color = $white })
        $mastLines.Add(@{ Text = ($title.En + ' - ' + $sub.En); Size = 12; Color = $white })
        Add-PMWordCellLines -Cell $mcell -Lines $mastLines

        # ---------- meta grid ----------
        $metaRows = New-Object System.Collections.Generic.List[object]
        $metaRows.Add(@{ K = 'ui.hostname'; Th = $Meta.Hostname; En = '' })
        if (-not [string]::IsNullOrWhiteSpace([string]$cfg.Report.SystemNameTh)) {
            $metaRows.Add(@{ K = 'ui.systemName'; Th = [string]$cfg.Report.SystemNameTh; En = [string]$cfg.Report.SystemNameEn })
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$cfg.Organization.DepartmentTh)) {
            $metaRows.Add(@{ K = 'ui.department'; Th = [string]$cfg.Organization.DepartmentTh; En = [string]$cfg.Organization.DepartmentEn })
        }
        $metaRows.Add(@{ K = 'ui.checkDate'; Th = $date.Th; En = $date.En })
        if (-not [string]::IsNullOrWhiteSpace([string]$cfg.Report.PreparedBy)) {
            $metaRows.Add(@{ K = 'ui.preparedBy'; Th = [string]$cfg.Report.PreparedBy; En = '' })
        }

        $r = $Doc.Content; $r.Collapse(0) | Out-Null
        $r.InsertParagraphAfter()
        $r = $Doc.Content; $r.Collapse(0) | Out-Null
        $mg = $doc.Tables.Add($r, $metaRows.Count, 2)
        $mg.Borders.Enable = -1
        $mg.Rows.AllowBreakAcrossPages = 0
        $mg.Columns.Item(1).Width = 130
        for ($i = 0; $i -lt $metaRows.Count; $i++) {
            $m = $metaRows[$i]
            $lab = Get-PMText -Key $m.K
            Set-PMWordCellBilingual -Cell $mg.Cell($i + 1, 1) -Th $lab.Th -En $lab.En -Bold
            $mg.Cell($i + 1, 1).Shading.BackgroundPatternColor = $Script:PMDocxHead
            Set-PMWordCellBilingual -Cell $mg.Cell($i + 1, 2) -Th ([string]$m.Th) -En ([string]$m.En)
        }

        # ---------- executive summary ----------
        Add-PMWordParagraph -Doc $doc -Style (-2) -Text ((Get-PMText -Key 'ui.summaryHeading').Th + ' / ' + (Get-PMText -Key 'ui.summaryHeading').En) | Out-Null

        $r = $Doc.Content; $r.Collapse(0) | Out-Null
        $r.InsertParagraphAfter()
        $r = $Doc.Content; $r.Collapse(0) | Out-Null
        $shown = @('CRIT', 'WARN', 'OK', 'INFO', 'ERROR') | Where-Object { $counts[$_] -gt 0 -or ($_ -ne 'ERROR' -and $_ -ne 'INFO') }
        $tiles = $doc.Tables.Add($r, 2, $shown.Count)
        $tiles.Borders.Enable = -1
        $tiles.Rows.AllowBreakAcrossPages = 0
        for ($i = 0; $i -lt $shown.Count; $i++) {
            $s = $shown[$i]
            $lab = Get-PMText -Key "status.$s"
            $sc = Get-PMDocxStatusColor -Status $s
            Set-PMWordCellBilingual -Cell $tiles.Cell(1, $i + 1) -Th ([string]$counts[$s]) -Bold
            $tiles.Cell(1, $i + 1).Range.Font.Size = 18
            $tiles.Cell(1, $i + 1).Range.Font.Color = $sc.FG
            $tiles.Cell(1, $i + 1).Shading.BackgroundPatternColor = $sc.BG
            Set-PMWordCellBilingual -Cell $tiles.Cell(2, $i + 1) -Th $lab.Th -En $lab.En
            $tiles.Cell(2, $i + 1).Range.Font.Size = 9
        }

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
        Add-PMWordParagraph -Doc $doc -Text $verdict.Th -Bold -Color $Script:PMDocxNavy | Out-Null
        Add-PMWordParagraph -Doc $doc -Text $verdict.En -Size 9 -Color $Script:PMDocxMuted -Tight | Out-Null

        if (-not $Meta.IsAdmin) {
            $w = Get-PMText -Key 'ui.notAdminWarning'
            $sc = Get-PMDocxStatusColor -Status 'WARN'
            Add-PMWordParagraph -Doc $doc -Text ($w.Th + ' / ' + $w.En) -Color $sc.FG | Out-Null
        }

        # ---------- overview ----------
        Add-PMWordParagraph -Doc $doc -Style (-2) -Text ((Get-PMText -Key 'ui.overviewHeading').Th + ' / ' + (Get-PMText -Key 'ui.overviewHeading').En) | Out-Null

        $r = $Doc.Content; $r.Collapse(0) | Out-Null
        $r.InsertParagraphAfter()
        $r = $Doc.Content; $r.Collapse(0) | Out-Null
        $ov = $doc.Tables.Add($r, $Results.Count + 1, 3)
        $ov.Borders.Enable = -1
        $ov.Rows.AllowBreakAcrossPages = 0
        $ov.Range.Font.Size = 10
        foreach ($pair in @(@(1, 'ui.col.topic'), @(2, 'ui.col.status'), @(3, 'ui.col.summary'))) {
            $c = Get-PMText -Key $pair[1]
            Set-PMWordCellBilingual -Cell $ov.Cell(1, $pair[0]) -Th $c.Th -En $c.En -Bold
        }
        $ov.Rows.Item(1).Shading.BackgroundPatternColor = $Script:PMDocxHead
        $ov.Rows.Item(1).HeadingFormat = -1
        for ($i = 0; $i -lt $Results.Count; $i++) {
            $res = $Results[$i]
            $wRow = $i + 2
            Set-PMWordCellBilingual -Cell $ov.Cell($wRow, 1) -Th $res.TitleTh -En $res.TitleEn
            $t = Get-PMText -Key "status.$($res.Status)"
            $sc = Get-PMDocxStatusColor -Status $res.Status
            Set-PMWordCellBilingual -Cell $ov.Cell($wRow, 2) -Th $t.Th -En $t.En
            $ov.Cell($wRow, 2).Range.Font.Color = $sc.FG
            if ($res.Status -in @('WARN', 'CRIT', 'ERROR')) { $ov.Rows.Item($wRow).Shading.BackgroundPatternColor = $sc.BG }
            Set-PMWordCellBilingual -Cell $ov.Cell($wRow, 3) -Th $res.SummaryTh -En $res.SummaryEn
        }

        # ---------- recommendations ----------
        # Same gate as the HTML report: Report.ShowRecommendations in
        # settings.json. Findings are always present in PM-Data.json
        # regardless, so nothing is lost by the gate being off.
        if (Get-PMSetting -Path 'Report.ShowRecommendations' -Default $false) {
            $sevRank = @{ 'CRIT' = 0; 'ERROR' = 1; 'WARN' = 2; 'INFO' = 3 }
            $findings = New-Object System.Collections.Generic.List[object]
            foreach ($res in $Results) {
                foreach ($f in $res.Findings) {
                    if ($null -eq $f) { continue }
                    $findings.Add([pscustomobject]@{
                        Severity = $f.Severity; Th = $f.Th; En = $f.En
                        TitleTh = $res.TitleTh; Rank = $sevRank[[string]$f.Severity]
                    })
                }
            }
            $sorted = @($findings | Sort-Object Rank, TitleTh)

            Add-PMWordParagraph -Doc $doc -Style (-2) -Text ((Get-PMText -Key 'ui.recoHeading').Th + ' / ' + (Get-PMText -Key 'ui.recoHeading').En) | Out-Null
            if ($sorted.Count -eq 0) {
                $n = Get-PMText -Key 'ui.recoNone'
                Add-PMWordParagraph -Doc $doc -Text ($n.Th + ' / ' + $n.En) -Color $Script:PMDocxMuted | Out-Null
            }
            else {
                for ($i = 0; $i -lt $sorted.Count; $i++) {
                    $f = $sorted[$i]
                    $t = Get-PMText -Key "status.$($f.Severity)"
                    $sc = Get-PMDocxStatusColor -Status $f.Severity
                    Add-PMWordParagraph -Doc $doc -Text ("$($i + 1). [$($t.Th)] $($f.TitleTh)") -Bold -Color $sc.FG | Out-Null
                    Add-PMWordParagraph -Doc $doc -Text $f.Th -Size 10 | Out-Null
                    Add-PMWordParagraph -Doc $doc -Text $f.En -Size 9 -Color $Script:PMDocxFaint -Tight | Out-Null
                }
            }
        }

        # ---------- details ----------
        Add-PMWordParagraph -Doc $doc -Style (-2) -Text ((Get-PMText -Key 'ui.detailHeading').Th + ' / ' + (Get-PMText -Key 'ui.detailHeading').En) | Out-Null

        foreach ($res in $Results) {
            $t = Get-PMText -Key "status.$($res.Status)"
            $sc = Get-PMDocxStatusColor -Status $res.Status
            Add-PMWordParagraph -Doc $doc -Style (-3) -Text ("$($res.TitleTh) ($($res.TitleEn)) - $($t.Th) / $($t.En)") | Out-Null
            $doc.Content.Paragraphs.Last.Range.Font.Color = $sc.FG

            if (-not [string]::IsNullOrWhiteSpace($res.SummaryTh)) {
                Add-PMWordParagraph -Doc $doc -Text $res.SummaryTh -Size 10 | Out-Null
                if (-not [string]::IsNullOrWhiteSpace($res.SummaryEn)) {
                    Add-PMWordParagraph -Doc $doc -Text $res.SummaryEn -Size 9 -Color $Script:PMDocxFaint -Tight | Out-Null
                }
            }

            if ($res.Chart) {
                $note = Get-PMText -Key 'ui.docxChartNote'
                Add-PMWordParagraph -Doc $doc -Text $note.Th -Size 9 -Color $Script:PMDocxFaint | Out-Null
            }

            if ($res.Rows -and $res.Rows.Count -gt 0) {
                New-PMDocxResultTable -Doc $doc -Result $res
            }
            else {
                $n = Get-PMText -Key 'ui.noRows'
                Add-PMWordParagraph -Doc $doc -Text ($n.Th + ' / ' + $n.En) -Size 9 -Color $Script:PMDocxFaint | Out-Null
            }
        }

        # ---------- footer ----------
        $foot = Get-PMText -Key 'ui.generatedFooter' -Values @($Meta.ToolVersion, $date.Th, $Meta.DurationSec)
        Add-PMWordParagraph -Doc $doc -Text $foot.Th -Size 8.5 -Color $Script:PMDocxFaint | Out-Null

        $outDir = Split-Path -Parent $Path
        if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

        # wdFormatDocumentDefault = 16 (.docx)
        $doc.SaveAs([ref]$Path, [ref]16)
    }
    finally {
        if ($doc)  { $doc.Close([ref]0); [Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null }
        if ($word) { $word.Quit(); [Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
