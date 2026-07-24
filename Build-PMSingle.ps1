<#
    PMtools - build the single-file distributable.

    Packs Lib\, Checks\, Config\ and the three entry scripts into one
    self-extracting file, so the tool can be handed to a server as a single
    attachment instead of a folder of thirty.

        .\Build-PMSingle.ps1
        .\Build-PMSingle.ps1 -Format Both

    The built file carries a base64 zip of the source tree. At run time it
    unpacks into a temporary folder, runs from there, and deletes the folder
    again - so the shipped artefact is a container, never a rewrite of the
    sources. Nothing here transforms the scripts, which is the whole point:
    the single file behaves exactly like the folder it was built from, and
    adding a check needs no change to this builder.

    ASCII-only by design, like every other .ps1 here - see Core.ps1. The Thai
    in Config\*.json survives regardless, because base64 is byte-exact.
#>
[CmdletBinding()]
param(
    # Where the built file is written. Default: Dist\ beside this script.
    [string]$OutDir,

    # Cmd - one .cmd: double-clickable, self-elevating, and immune to the
    #       machine's execution policy. This is the one to hand an operator.
    # Ps1 - one .ps1: for scheduled tasks and pipelines that start
    #       powershell.exe themselves and do not want the elevation prompt.
    [ValidateSet('Cmd', 'Ps1', 'Both')][string]$Format = 'Cmd',

    # Base file name without extension. Default: PMtools-<ToolVersion>.
    [string]$Name
)

$ErrorActionPreference = 'Stop'

$PMRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($PMRoot)) { $PMRoot = $PSScriptRoot }

if ([string]::IsNullOrWhiteSpace($OutDir)) { $OutDir = Join-Path $PMRoot 'Dist' }

# ---------------------------------------------------------------------
# What goes in the bundle
#   README.md and docs\Manual.html stay out: the payload is unpacked to a
#   temporary folder and deleted again, so documentation shipped inside it
#   would never be seen by anybody. Those travel separately.
# ---------------------------------------------------------------------

$bundleFiles = @(
    'Start-PMCheck.ps1'
    'Start-PMMonitor.ps1'
    'Show-PMMenu.ps1'
)

$bundleDirs = @(
    @{ Name = 'Lib';    Filter = '*.ps1'  }
    @{ Name = 'Checks'; Filter = '*.ps1'  }
    @{ Name = 'Config'; Filter = '*.json' }
)

# Report.Docx.ps1 needs Microsoft Word and runs on an administrator's own
# workstation via Export-PMDocxReport.ps1 - it has no part in the read-only
# server assessment this single file exists to carry, so it is left out
# rather than riding along unused.
$bundleExclude = @('Report.Docx.ps1')

# ---------------------------------------------------------------------
# Stage, zip, encode
# ---------------------------------------------------------------------

Write-Host ''
Write-Host 'PMtools - building the single-file distributable' -ForegroundColor Cyan
Write-Host ''

$version = '1.7.5'
try {
    $settings = Get-Content -LiteralPath (Join-Path $PMRoot 'Config\settings.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($settings.Report.ToolVersion) { $version = [string]$settings.Report.ToolVersion }
}
catch { Write-Warning "Could not read Report.ToolVersion from settings.json; using $version." }

if ([string]::IsNullOrWhiteSpace($Name)) { $Name = "PMtools-$version" }

$stage    = Join-Path ([System.IO.Path]::GetTempPath()) ('PMbuild-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
$zipPath  = "$stage.zip"
$staged   = 0
$zipBytes = $null

New-Item -ItemType Directory -Path $stage -Force | Out-Null
try {
    foreach ($file in $bundleFiles) {
        $src = Join-Path $PMRoot $file
        if (-not (Test-Path -LiteralPath $src)) { throw "Missing required file: $file" }
        Copy-Item -LiteralPath $src -Destination (Join-Path $stage $file) -Force
        $staged++
    }
    Write-Host ("  {0,-8} {1,3} file(s)" -f 'root', $bundleFiles.Count)

    foreach ($dir in $bundleDirs) {
        $src = Join-Path $PMRoot $dir.Name
        if (-not (Test-Path -LiteralPath $src)) { throw "Missing required folder: $($dir.Name)" }

        $dest = Join-Path $stage $dir.Name
        New-Item -ItemType Directory -Path $dest -Force | Out-Null

        $items = @(Get-ChildItem -LiteralPath $src -Filter $dir.Filter -File |
                   Where-Object { $bundleExclude -notcontains $_.Name } | Sort-Object Name)
        if ($items.Count -eq 0) { throw "Folder $($dir.Name) holds no $($dir.Filter) files." }

        foreach ($item in $items) {
            Copy-Item -LiteralPath $item.FullName -Destination (Join-Path $dest $item.Name) -Force
            $staged++
        }
        Write-Host ("  {0,-8} {1,3} file(s)" -f $dir.Name, $items.Count)
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $stage, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)

    $zipBytes = [System.IO.File]::ReadAllBytes($zipPath)
}
finally {
    Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
}

# Wrapped so no single line runs long enough to trouble an editor, a diff or
# a mail gateway that reflows what it thinks is a text attachment.
$b64   = [Convert]::ToBase64String($zipBytes)
$lines = New-Object System.Collections.Generic.List[string]
for ($i = 0; $i -lt $b64.Length; $i += 120) {
    $lines.Add($b64.Substring($i, [Math]::Min(120, $b64.Length - $i)))
}

# Assembled here rather than inside the template below, because a here-string
# cannot contain the delimiters of the here-string that would hold it.
$payloadBlock = "@'" + "`r`n" + ($lines -join "`r`n") + "`r`n" + "'@"

Write-Host ''
Write-Host ("  staged {0} file(s), {1:N0} KB compressed" -f $staged, ($zipBytes.Length / 1KB))

# ---------------------------------------------------------------------
# The bootstrap that ships inside the built file
#   One template shared by both formats. It locates its own folder, unpacks
#   the payload, hands control to the ordinary entry scripts, and removes the
#   temporary folder on the way out - including when a check throws, which is
#   why the run sits in a try/finally.
# ---------------------------------------------------------------------

$bootstrap = @'
param(
    # Write Config\ next to this file so thresholds and the organisation name
    # can be edited. Afterwards any *.json found there overrides the bundled
    # copy file by file, so a lone settings.json is enough - i18n.json does
    # not have to be kept alongside it.
    [switch]$ExtractConfig,

    # Leave the unpacked folder behind and print where it is. For diagnosing
    # a failure that only reproduces from the single file.
    [switch]$KeepFiles
)

# Anything not named above goes straight through to Start-PMCheck.ps1:
#   -Only DISK,CERT   -Skip WU   -MonitorMinutes 30   -OpenReport
# Captured here because $args inside a function is that function's own.
# Left untyped on purpose: -Only DISK,CERT arrives as one nested array, and
# declaring [string[]] would flatten it into two separate arguments.
$PMForward = @($args)

$ErrorActionPreference = 'Stop'
$Script:PMExitCode = 0

# Splatting an ARRAY binds positionally - '-OutputRoot' would arrive as the
# value of the first parameter rather than naming it, which is exactly the
# fault this function exists to avoid. Only a hashtable splats by name, so
# the forwarded tokens are folded into one here: '-Name value' becomes an
# entry, a '-Name' with no value behind it becomes a switch.
function ConvertTo-PMArgumentTable {
    param([object[]]$Tokens)

    # powershell.exe -File hands every argument over as a plain string, so
    # '-Only DISK,CERT' arrives as one token here, while the -Command form the
    # .cmd wrapper uses splits it into an array first. These are the two
    # parameters Start-PMCheck.ps1 declares as string[]; put them back into
    # lists so both ways of launching behave the same.
    $listParams = @('Only', 'Skip')

    $table = @{}
    $i = 0
    while ($i -lt $Tokens.Count) {
        $name = "$($Tokens[$i])"
        if ($name -notlike '-*') {
            throw "Unexpected argument '$name'. Use named parameters, for example: -Only DISK,CERT -OpenReport"
        }
        $name = $name.TrimStart('-')

        # A value is whatever follows that is not itself a parameter name.
        if (($i + 1) -lt $Tokens.Count -and "$($Tokens[$i + 1])" -notlike '-*') {
            $value = $Tokens[$i + 1]
            if ($listParams -contains $name -and $value -is [string]) {
                $value = @($value.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            }
            $table[$name] = $value
            $i += 2
        }
        else {
            $table[$name] = $true
            $i++
        }
    }
    return $table
}

function Invoke-PMSingleFile {
    param(
        [Parameter(Mandatory)][string]$Base64,
        [object[]]$Forward
    )

    # PM_SELF is set by the .cmd wrapper before it hands over; $PSCommandPath
    # covers the .ps1 build, which is launched as a file in the ordinary way.
    $self = $env:PM_SELF
    if ([string]::IsNullOrWhiteSpace($self)) { $self = $PSCommandPath }

    if ([string]::IsNullOrWhiteSpace($self)) { $selfDir = (Get-Location).Path }
    else                                     { $selfDir = Split-Path -Parent $self }

    $work = Join-Path ([System.IO.Path]::GetTempPath()) ('PMtools-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Path $work -Force | Out-Null

    try {
        $zip = Join-Path $work '_payload.zip'
        [System.IO.File]::WriteAllBytes($zip, [Convert]::FromBase64String(($Base64 -replace '[^A-Za-z0-9+/=]', '')))
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $work)
        Remove-Item -LiteralPath $zip -Force

        $side = Join-Path $selfDir 'Config'

        if ($ExtractConfig) {
            New-Item -ItemType Directory -Path $side -Force | Out-Null
            Copy-Item -Path (Join-Path $work 'Config\*') -Destination $side -Force
            Write-Host ''
            Write-Host ("Config written to {0}" -f $side) -ForegroundColor Green
            Write-Host 'Edit settings.json there. Later runs pick it up on their own.'
            Write-Host ''
            return
        }

        # Overrides are copied one file at a time, so a partial Config\ left
        # beside the executable cannot leave the bundled set incomplete -
        # Core.ps1 throws outright if settings.json or i18n.json is missing.
        if (Test-Path -LiteralPath $side) {
            $over = @(Get-ChildItem -LiteralPath $side -Filter '*.json' -File -ErrorAction SilentlyContinue)
            if ($over.Count -gt 0) {
                foreach ($f in $over) {
                    Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $work 'Config') -Force
                }
                Write-Host ("Config override: {0} file(s) from {1}" -f $over.Count, $side) -ForegroundColor Cyan
            }
        }

        # The report has to outlive the temporary folder, so it is written
        # straight into the folder holding this file: drop the executable in a
        # folder, run it, and the <HOSTNAME>_<timestamp>\ folder appears right
        # next to it with nothing in between.
        #
        # Probed with a real file rather than New-Item on the folder itself,
        # which succeeds on a directory that already exists whether or not it
        # can be written to - and a share or a locked-down USB stick has to be
        # caught here, not after every check has already run.
        $outRoot = $selfDir
        try {
            $probe = Join-Path $outRoot ('pmwrite-' + [Guid]::NewGuid().ToString('N').Substring(0, 8) + '.tmp')
            New-Item -ItemType File -Path $probe -Force -ErrorAction Stop | Out-Null
            Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        }
        catch {
            $outRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'PM-Output'
            New-Item -ItemType Directory -Path $outRoot -Force | Out-Null
            Write-Warning ("{0} is not writable. Writing reports to {1} instead." -f $selfDir, $outRoot)
        }

        if (@($Forward).Count -gt 0) {
            $pass = ConvertTo-PMArgumentTable -Tokens $Forward
            # An -OutputRoot from the caller wins over the default above.
            if (-not $pass.ContainsKey('OutputRoot')) { $pass['OutputRoot'] = $outRoot }
            & (Join-Path $work 'Start-PMCheck.ps1') @pass
        }
        else {
            & (Join-Path $work 'Show-PMMenu.ps1') -OutputRoot $outRoot
        }
        if ($null -ne $LASTEXITCODE) { $Script:PMExitCode = $LASTEXITCODE }
    }
    finally {
        if ($KeepFiles) {
            Write-Host ''
            Write-Host ("Unpacked files kept at {0}" -f $work) -ForegroundColor Yellow
        }
        else {
            Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# --- payload: a zip of Lib\, Checks\, Config\ and the entry scripts ---------
$PMPayload = @@PAYLOAD_BLOCK@@

Invoke-PMSingleFile -Base64 $PMPayload -Forward $PMForward
exit $Script:PMExitCode
'@

$built = (Get-Date).ToString('yyyy-MM-dd HH:mm')

$psHeader = @'
# ===================================================================
#  PMtools @@VERSION@@ - single file, built @@BUILT@@
#
#  Generated by Build-PMSingle.ps1 from the PMtools source tree.
#  Do not edit this file: change the source and build it again.
#
#  Read-only, like the folder it was built from. It unpacks itself into
#  a temporary folder, reads the machine, writes the report folder next
#  to this file, and removes the temporary folder again.
#
#      powershell -ExecutionPolicy Bypass -File PMtools.ps1
#      powershell -ExecutionPolicy Bypass -File PMtools.ps1 -Only DISK,CERT
# ===================================================================

'@

$psText = ($psHeader + $bootstrap).
    Replace('@@PAYLOAD_BLOCK@@', $payloadBlock).
    Replace('@@VERSION@@', $version).
    Replace('@@BUILT@@', $built)

# ---------------------------------------------------------------------
# The .cmd wrapper
#   cmd.exe reads a batch file line by line as it runs it, so the PowerShell
#   half past `exit /b` is never parsed - it is only bytes on disk that the
#   batch half reads back out of its own file. That is what makes a single
#   double-clickable file possible with no execution policy to argue with.
# ---------------------------------------------------------------------

$cmdHeader = @'
@echo off
setlocal enableextensions
rem ===================================================================
rem  PMtools @@VERSION@@ - single file, built @@BUILT@@
rem
rem  Copy this one file to the server. Nothing else needs to go with it.
rem
rem    double-click                  menu; asks for Administrator first
rem    PMtools.cmd -Only DISK,CERT   run named checks, then exit
rem    PMtools.cmd -ExtractConfig    write Config\ here so it can be edited
rem
rem  The report lands in <HOSTNAME>_<date-time>\ right beside this file.
rem  Nothing on the server is changed - every check only reads.
rem
rem  Past the exit below is PowerShell and a base64 payload. cmd.exe
rem  never reads that far.
rem ===================================================================

set "PM_SELF=%~f0"

net session >nul 2>&1
if not errorlevel 1 goto :elevated

echo Requesting Administrator privileges...
if "%~1"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%PM_SELF%' -Verb RunAs"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%PM_SELF%' -ArgumentList '%*' -Verb RunAs"
)
exit /b 0

:elevated
rem The marker is spliced back together so IndexOf cannot match this very line.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:PM_SELF); $m='#PM'+'TOOLS_PAYLOAD#'; & ([scriptblock]::Create($t.Substring($t.IndexOf($m))))" %*
set PMEXIT=%errorlevel%

rem Hold the window open only for a double-click. A caller that passed
rem arguments wants the exit code back, not a keypress.
if not "%~1"=="" goto :done
echo.
if "%PMEXIT%"=="0" echo Result: no action items.
if "%PMEXIT%"=="1" echo Result: items requiring monitoring were found - see the report.
if "%PMEXIT%"=="2" echo Result: critical items were found - see the report.
if %PMEXIT% GTR 2 echo Result: the tool did not finish. Review the messages above.
echo.
echo Press any key to close this window.
pause >nul

:done
endlocal & exit /b %PMEXIT%

#PMTOOLS_PAYLOAD#
'@

$cmdText = $cmdHeader.Replace('@@VERSION@@', $version).Replace('@@BUILT@@', $built) + $psText

# ---------------------------------------------------------------------
# Write
# ---------------------------------------------------------------------

function Write-PMArtifact {
    param([string]$Path, [string]$Text)

    # ASCII with no BOM. A BOM at the top of a .cmd is echoed as garbage by
    # cmd.exe, and PowerShell 5.1 reads a BOM-less .ps1 as ANSI - harmless
    # only because every byte written here is ASCII, base64 included.
    $normalized = ($Text -replace "`r`n", "`n") -replace "`n", "`r`n"
    [System.IO.File]::WriteAllText($Path, $normalized, (New-Object System.Text.ASCIIEncoding))

    Write-Host ("  {0}  ({1:N0} KB)" -f $Path, ((Get-Item -LiteralPath $Path).Length / 1KB)) -ForegroundColor Green
}

if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

Write-Host ''
Write-Host 'Built:' -ForegroundColor Cyan

if ($Format -eq 'Cmd' -or $Format -eq 'Both') {
    Write-PMArtifact -Path (Join-Path $OutDir "$Name.cmd") -Text $cmdText
}
if ($Format -eq 'Ps1' -or $Format -eq 'Both') {
    Write-PMArtifact -Path (Join-Path $OutDir "$Name.ps1") -Text $psText
}

Write-Host ''
Write-Host 'Copy that file to the server and run it. Nothing else goes with it.'
Write-Host ''
