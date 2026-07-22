<#
    PMtools - Word (.docx) export.

    Reads a PM-Data.json already produced by Start-PMCheck.ps1 and writes a
    bilingual Thai/English Word document alongside it. Meant to run on an
    administrator's own workstation, not the server that was assessed: it
    needs Microsoft Word installed, and it needs nothing from the server
    beyond the one JSON file - copy PM-Data.json off the server (email, a
    share, a USB stick) and run this here.

    Does not need Administrator rights and does not elevate: it only reads a
    file already on disk and drives Word through COM, neither of which
    touches machine state.

        .\Export-PMDocxReport.ps1 -DataPath D:\reports\PM-Data.json
        .\Export-PMDocxReport.ps1 -DataPath D:\reports\PM-Data.json -OpenReport
        .\Export-PMDocxReport.ps1 -DataPath D:\reports\PM-Data.json -OutputPath D:\reports\Final.docx

    PM-Data.json deliberately does not carry Organization/Department/System/
    PreparedBy - those come from Config\settings.json instead, read fresh
    here exactly as Start-PMCheck.ps1 reads them when it builds the HTML
    report. Edit Config\settings.json beside this script to change them;
    see README.md.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DataPath,

    [string]$ConfigDir,

    # Default: PM-Report.docx next to the input JSON.
    [string]$OutputPath,

    [switch]$OpenReport
)

$ErrorActionPreference = 'Stop'

# See the fuller note in Start-PMCheck.ps1: $PSScriptRoot is empty while
# param() defaults are evaluated when a [CmdletBinding()] script runs under
# `powershell.exe -File`, so the folder is resolved here in the body instead.
$PMRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($PMRoot)) { $PMRoot = $PSScriptRoot }

if ([string]::IsNullOrWhiteSpace($ConfigDir)) { $ConfigDir = Join-Path $PMRoot 'Config' }

if (-not (Test-Path -LiteralPath $DataPath)) { throw "Data file not found: $DataPath" }
$DataPath = (Resolve-Path -LiteralPath $DataPath).Path

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path (Split-Path -Parent $DataPath) 'PM-Report.docx'
}
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)

# --- load the shared contract and the Word renderer -------------------------
. (Join-Path $PMRoot 'Lib\Core.ps1')
. (Join-Path $PMRoot 'Lib\Report.Docx.ps1')

Initialize-PMCore -ConfigDir $ConfigDir

Write-PMLog ''
Write-PMLog 'PMtools - Word report export' -Level Step
Write-PMLog ("Data   : {0}" -f $DataPath)
Write-PMLog ''

# --- read the assessment data -----------------------------------------------
# PowerShell 5.1's ConvertFrom-Json turns the "/Date(...)/ " form ConvertTo-
# Json wrote back into a real [datetime] on the way in, so GeneratedAt needs
# no special handling here - confirmed by round-tripping an actual PM-Data.
# json rather than assumed from documentation.
$payload = Get-Content -LiteralPath $DataPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $payload.Meta -or -not $payload.Results) {
    throw "This does not look like a PM-Data.json file (missing Meta or Results): $DataPath"
}

# Config travels with this script, not with the JSON - New-PMDocxReport reads
# Organization/Department/SystemName/PreparedBy from it the same way
# Start-PMCheck.ps1 does for the HTML report, so edit Config\settings.json
# beside this script to change what shows up.
$meta = [pscustomobject]@{
    Hostname    = $payload.Meta.Hostname
    GeneratedAt = $payload.Meta.GeneratedAt
    DurationSec = $payload.Meta.DurationSec
    ToolVersion = $payload.Meta.ToolVersion
    IsAdmin     = $payload.Meta.IsAdmin
    Config      = Get-PMConfig
}

$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    New-PMDocxReport -Results @($payload.Results) -Meta $meta -Path $OutputPath
}
catch {
    Write-PMLog ("Failed: {0}" -f $_.Exception.Message) -Level Bad
    throw
}
$sw.Stop()

Write-PMLog ("Report : {0}" -f $OutputPath) -Level Good
Write-PMLog ("Done in {0:N1} s" -f ($sw.ElapsedMilliseconds / 1000))
Write-PMLog ''

if ($OpenReport) { Start-Process $OutputPath }
