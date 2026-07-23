# =====================================================================
#  PMtools - ArcGIS.ps1
#  Connection, authentication and REST plumbing for the ArcGIS Server
#  Administrator API. Shared by every Checks\A*-ArcGIS*.ps1 check and by
#  the connection menu in Show-PMMenu.ps1.
#
#  ASCII-only, same rule as every other .ps1 here - see Core.ps1.
#
#  READ-ONLY, deliberately. Only two verbs are used against the site:
#    - GET  for everything that reads state.
#    - POST for generateToken, data/findItems and data/validateDataItem
#      ONLY. All three are POST because the Admin API requires it for
#      these operations (a token, or a JSON item description, must travel
#      in the request body), NOT because they change anything:
#      generateToken issues a short-lived token, data/findItems only lists
#      what is already registered, and data/validateDataItem only retests
#      a connection the site already has stored credentials for, using
#      those credentials - see Checks\A3-ArcGISData.ps1 (AGSDATA) for how
#      it is used and, just as important, what it is careful never to
#      keep: findItems returns each item's connection string, encrypted
#      password included, and AGSDATA never lets that leave memory.
#      Nothing here writes to the site.
#
#  What is deliberately NOT implemented: creating a usage report. Reading
#  one that already exists (Checks\A2-ArcGISUsage.ps1, AGSUSAGE) is a plain
#  GET against /usagereports and /usagereports/<name>/data, and is fine -
#  but querying a report that has never been created is not possible
#  without creating it first, and creating one is a write. Service instance
#  statistics from /admin/services/<folder>/report give a related but
#  different answer (is anything saturated, per service?) through a plain
#  GET that never depends on a report existing - see Checks\A1-ArcGISServices.ps1.
# =====================================================================

# ---------------------------------------------------------------------
# Transport
# ---------------------------------------------------------------------

# PowerShell 5.1 negotiates SSL3/TLS1.0 by default on .NET 4.x, which every
# supported ArcGIS Server release now refuses - the failure surfaces as a
# bare "The underlying connection was closed", with nothing pointing at TLS.
# Raising it once per process costs nothing and removes that whole class of
# confusing failure.
function Initialize-PMArcGISTransport {
    try {
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
    catch { }
}

# ArcGIS Server ships with a self-signed certificate on port 6443, and most
# sites never replace it on the direct port even when the Web Adaptor in
# front has a real one. Validating it would make this check unusable on the
# majority of real installations, so it is bypassed by default - the
# operator typed the hostname themselves and the call goes to their own
# server.
#
# ServerCertificateValidationCallback is process-wide with no scoped
# alternative on PS 5.1, so the previous value is captured and restored by
# Restore-PMArcGISCertificatePolicy rather than left switched off for
# whatever runs after this check.
$Script:PMArcGISPriorCertCallback = $null
$Script:PMArcGISCertBypassed      = $false

function Disable-PMArcGISCertificateCheck {
    if ($Script:PMArcGISCertBypassed) { return }
    $Script:PMArcGISPriorCertCallback = [Net.ServicePointManager]::ServerCertificateValidationCallback
    [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    $Script:PMArcGISCertBypassed = $true
}

function Restore-PMArcGISCertificatePolicy {
    if (-not $Script:PMArcGISCertBypassed) { return }
    [Net.ServicePointManager]::ServerCertificateValidationCallback = $Script:PMArcGISPriorCertCallback
    $Script:PMArcGISPriorCertCallback = $null
    $Script:PMArcGISCertBypassed      = $false
}

# ---------------------------------------------------------------------
# URL handling
# ---------------------------------------------------------------------

# Turns whatever the operator typed into the site root that /admin hangs
# off. The three forms people actually paste are all accepted:
#
#   https://gis.example.go.th/server          -> .../server         (Web Adaptor)
#   https://gis.example.go.th:6443            -> .../arcgis         (direct, path assumed)
#   https://gis.example.go.th:6443/arcgis     -> .../arcgis         (direct, path given)
#
# A trailing /admin is tolerated and stripped, because that is the URL
# shown in the browser address bar when someone is already looking at the
# Administrator Directory and copies it.
function Get-PMArcGISRoot {
    param([Parameter(Mandatory)][string]$Url)

    $u = $Url.Trim()
    if ([string]::IsNullOrWhiteSpace($u)) { throw 'The ArcGIS Server URL is empty.' }

    if ($u -notmatch '^https?://') { $u = 'https://' + $u }
    $u = $u.TrimEnd('/')
    $u = $u -replace '(?i)/admin/?$', ''
    $u = $u.TrimEnd('/')

    $uri = $null
    if (-not [Uri]::TryCreate($u, [UriKind]::Absolute, [ref]$uri)) {
        throw "Not a valid URL: $Url"
    }

    # Host with no path at all: the default site name is the only sensible
    # guess, and it is right on a default install.
    if ([string]::IsNullOrWhiteSpace($uri.AbsolutePath) -or $uri.AbsolutePath -eq '/') {
        $u = $u + '/arcgis'
    }

    return $u
}

function Get-PMArcGISAdminUrl {
    param([Parameter(Mandatory)][string]$Root, [string]$Path = '')
    $p = $Path.TrimStart('/')
    if ([string]::IsNullOrWhiteSpace($p)) { return ($Root + '/admin') }
    return ($Root + '/admin/' + $p)
}

# ---------------------------------------------------------------------
# Stored connection
#   URL and username are ordinary settings. The password is encrypted with
#   DPAPI through ConvertFrom-SecureString, which binds the ciphertext to
#   this Windows account on this machine: copying the file to another
#   machine, or reading it as another user, yields nothing. That is the
#   strongest option available without adding a dependency, and it is why
#   the password is never put in settings.json - that file is committed to
#   git and is handed out by -ExtractConfig.
# ---------------------------------------------------------------------

# Stored under the Windows profile rather than in Config\, for two reasons.
#
# The single-file build unpacks itself into a temporary folder that is
# deleted the moment the run finishes, so a connection saved beside
# settings.json there would not survive to the next run - and the single
# file is the primary way this tool is handed out.
#
# It also matches the security boundary already in force: DPAPI ciphertext
# only opens for the account that wrote it, on the machine that wrote it,
# so the per-user profile is exactly the right scope for it to live in.
#
# $ConfigDir is still honoured when passed, which keeps the path injectable
# for testing.
function Get-PMArcGISConnectionPath {
    param([string]$ConfigDir)

    if (-not [string]::IsNullOrWhiteSpace($ConfigDir)) {
        return (Join-Path $ConfigDir 'arcgis-connection.xml')
    }

    $base = $env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($base)) { $base = $env:APPDATA }
    if ([string]::IsNullOrWhiteSpace($base)) { $base = [System.IO.Path]::GetTempPath() }

    return (Join-Path (Join-Path $base 'PMtools') 'arcgis-connection.xml')
}

function Save-PMArcGISConnection {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$Username,
        [Parameter(Mandatory)][System.Security.SecureString]$Password,
        [string]$ConfigDir
    )
    $path = Get-PMArcGISConnectionPath -ConfigDir $ConfigDir
    $dir  = Split-Path -Parent $path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    [pscustomobject]@{
        Url      = Get-PMArcGISRoot -Url $Url
        Username = $Username
        Password = ConvertFrom-SecureString -SecureString $Password
        SavedBy  = "$env:USERDOMAIN\$env:USERNAME"
        SavedOn  = $env:COMPUTERNAME
        SavedAt  = (Get-Date).ToString('s')
    } | Export-Clixml -LiteralPath $path -Force

    return $path
}

function Get-PMArcGISConnection {
    param([string]$ConfigDir)

    $path = Get-PMArcGISConnectionPath -ConfigDir $ConfigDir
    if (-not (Test-Path -LiteralPath $path)) { return $null }

    try { $saved = Import-Clixml -LiteralPath $path }
    catch { throw "The saved ArcGIS connection at $path could not be read: $($_.Exception.Message)" }

    $secure = $null
    if ($saved.Password) {
        try { $secure = ConvertTo-SecureString -String ([string]$saved.Password) }
        catch {
            # DPAPI refuses when the file was written by a different Windows
            # account or on a different machine. Say so plainly - the generic
            # "Key not valid for use in specified state" gives an operator
            # nothing to act on.
            throw ("The saved password cannot be decrypted here. It was saved by {0} on {1}; a DPAPI secret only opens for the same account on the same machine. Set the connection again from this account." -f $saved.SavedBy, $saved.SavedOn)
        }
    }

    return [pscustomobject]@{
        Url      = [string]$saved.Url
        Username = [string]$saved.Username
        Password = $secure
        SavedBy  = [string]$saved.SavedBy
        SavedOn  = [string]$saved.SavedOn
        SavedAt  = [string]$saved.SavedAt
        Path     = $path
    }
}

function Clear-PMArcGISConnection {
    param([string]$ConfigDir)
    $path = Get-PMArcGISConnectionPath -ConfigDir $ConfigDir
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force; return $true }
    return $false
}

# ---------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------

$Script:PMArcGISToken       = $null
$Script:PMArcGISTokenExpiry = [datetime]::MinValue
$Script:PMArcGISTokenRoot   = ''

function Clear-PMArcGISToken {
    $Script:PMArcGISToken       = $null
    $Script:PMArcGISTokenExpiry = [datetime]::MinValue
    $Script:PMArcGISTokenRoot   = ''
}

# Converts a SecureString back to plain text for exactly as long as the web
# request needs it, then zeroes the unmanaged buffer. The plain password is
# never stored in a variable that outlives this function, never written to
# PM-Data.json, and never logged.
function ConvertFrom-PMSecureString {
    param([Parameter(Mandatory)][System.Security.SecureString]$Secure)
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try   { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

function Get-PMArcGISToken {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$Username,
        [Parameter(Mandatory)][System.Security.SecureString]$Password,
        [int]$ExpirationMinutes = 60,
        [int]$TimeoutSec = 30,
        [switch]$Force
    )

    if (-not $Force -and $Script:PMArcGISToken -and $Script:PMArcGISTokenRoot -eq $Root -and
        (Get-Date) -lt $Script:PMArcGISTokenExpiry) {
        return $Script:PMArcGISToken
    }

    Initialize-PMArcGISTransport
    if (Get-PMSetting -Path 'ArcGIS.IgnoreCertificateErrors' -Default $true) { Disable-PMArcGISCertificateCheck }

    $url   = Get-PMArcGISAdminUrl -Root $Root -Path 'generateToken'
    $plain = ConvertFrom-PMSecureString -Secure $Password

    # client=requestip ties the token to the calling address, so a token
    # captured off the wire is useless from anywhere else.
    $body = @{
        username   = $Username
        password   = $plain
        client     = 'requestip'
        expiration = $ExpirationMinutes
        f          = 'json'
    }

    try {
        $resp = Invoke-RestMethod -Uri $url -Method Post -Body $body -TimeoutSec $TimeoutSec -ErrorAction Stop
    }
    catch {
        throw ("Could not reach {0}: {1}" -f $url, $_.Exception.Message)
    }
    finally {
        $plain = $null
        $body['password'] = $null
        [GC]::Collect()
    }

    # The Admin API answers HTTP 200 with an error object in the body rather
    # than an HTTP error status, so a bad password never lands in the catch
    # above - it has to be detected here.
    if ($resp.PSObject.Properties['error']) {
        $msg = $resp.error.message
        if ($resp.error.details) { $msg = $msg + ' - ' + ($resp.error.details -join '; ') }
        throw ("ArcGIS Server rejected the sign-in: {0}" -f $msg)
    }
    if (-not $resp.token) {
        throw 'ArcGIS Server returned no token and no error. Check that the URL points at an ArcGIS Server site.'
    }

    $Script:PMArcGISToken     = [string]$resp.token
    $Script:PMArcGISTokenRoot = $Root
    # Renew a minute early so a token cannot expire mid-run.
    $Script:PMArcGISTokenExpiry = (Get-Date).AddMinutes([math]::Max(1, $ExpirationMinutes - 1))

    return $Script:PMArcGISToken
}

# ---------------------------------------------------------------------
# Requests
# ---------------------------------------------------------------------

# One call against the Administrator API. GET by default; -Method Post only
# for the two read-only POST endpoints noted in this file's header.
function Invoke-PMArcGISAdmin {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Token,
        [hashtable]$Parameters,
        [ValidateSet('Get', 'Post')][string]$Method = 'Get',
        [int]$TimeoutSec = 60
    )

    Initialize-PMArcGISTransport
    if (Get-PMSetting -Path 'ArcGIS.IgnoreCertificateErrors' -Default $true) { Disable-PMArcGISCertificateCheck }

    # Deliberately not named $args: inside a function that is PowerShell's
    # automatic variable for unbound arguments, and shadowing it makes the
    # binding of any later call in this scope hard to reason about.
    $reqArgs = @{}
    if ($Parameters) { foreach ($k in $Parameters.Keys) { $reqArgs[$k] = $Parameters[$k] } }
    $reqArgs['f']     = 'json'
    $reqArgs['token'] = $Token

    $url = Get-PMArcGISAdminUrl -Root $Root -Path $Path

    try {
        if ($Method -eq 'Post') {
            $resp = Invoke-RestMethod -Uri $url -Method Post -Body $reqArgs -TimeoutSec $TimeoutSec -ErrorAction Stop
        }
        else {
            $query = ($reqArgs.GetEnumerator() | ForEach-Object {
                '{0}={1}' -f [Uri]::EscapeDataString([string]$_.Key), [Uri]::EscapeDataString([string]$_.Value)
            }) -join '&'
            $resp = Invoke-RestMethod -Uri ($url + '?' + $query) -Method Get -TimeoutSec $TimeoutSec -ErrorAction Stop
        }
    }
    catch {
        throw ("Request to {0} failed: {1}" -f $url, $_.Exception.Message)
    }

    if ($resp -and $resp.PSObject.Properties['error']) {
        $msg = $resp.error.message
        if ($resp.error.details) { $msg = $msg + ' - ' + ($resp.error.details -join '; ') }
        throw ("{0} returned an error: {1}" -f $Path, $msg)
    }

    return $resp
}

# ---------------------------------------------------------------------
# Session helper used by the checks
#   Resolves the stored connection and gets a token once, so a run with
#   several ArcGIS checks signs in a single time.
# ---------------------------------------------------------------------

function Get-PMArcGISSession {
    param([string]$ConfigDir)

    $conn = Get-PMArcGISConnection -ConfigDir $ConfigDir
    if ($null -eq $conn) {
        throw 'No ArcGIS Server connection is configured. Run the launcher and choose "ArcGIS Server connection" to set one.'
    }
    if ([string]::IsNullOrWhiteSpace($conn.Url) -or [string]::IsNullOrWhiteSpace($conn.Username) -or $null -eq $conn.Password) {
        throw 'The saved ArcGIS Server connection is incomplete. Set it again from the launcher.'
    }

    $expiry  = [int](Get-PMSetting -Path 'ArcGIS.TokenExpirationMinutes' -Default 60)
    $timeout = [int](Get-PMSetting -Path 'ArcGIS.TimeoutSeconds' -Default 30)

    $token = Get-PMArcGISToken -Root $conn.Url -Username $conn.Username -Password $conn.Password `
                               -ExpirationMinutes $expiry -TimeoutSec $timeout

    return [pscustomobject]@{
        Root       = $conn.Url
        Username   = $conn.Username
        Token      = $token
        TimeoutSec = [int](Get-PMSetting -Path 'ArcGIS.RequestTimeoutSeconds' -Default 60)
    }
}

# Probe used by the menu's "Test connection". Returns a result object
# instead of throwing, because the menu wants to print the reason rather
# than unwind.
function Test-PMArcGISConnection {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)][string]$Username,
        [Parameter(Mandatory)][System.Security.SecureString]$Password,
        [int]$TimeoutSec = 30
    )

    $result = [pscustomobject]@{
        Success     = $false
        Root        = ''
        Message     = ''
        Version     = ''
        SiteName    = ''
        MachineCount = 0
    }

    try {
        $result.Root = Get-PMArcGISRoot -Url $Url
        $token = Get-PMArcGISToken -Root $result.Root -Username $Username -Password $Password `
                                   -TimeoutSec $TimeoutSec -Force

        $info = Invoke-PMArcGISAdmin -Root $result.Root -Path 'info' -Token $token -TimeoutSec $TimeoutSec
        if ($info.currentversion) { $result.Version = [string]$info.currentversion }

        try {
            $machines = Invoke-PMArcGISAdmin -Root $result.Root -Path 'machines' -Token $token -TimeoutSec $TimeoutSec
            if ($machines.machines) { $result.MachineCount = @($machines.machines).Count }
        }
        catch {
            # An account that can sign in but cannot list machines still
            # proves the URL and password are right, which is what the test
            # is for. Report the sign-in as a success and let the check
            # itself surface the privilege problem later.
            $result.Message = 'Signed in, but this account cannot list machines (a read-only role may be too restrictive for some checks).'
        }

        $result.Success = $true
        if ([string]::IsNullOrWhiteSpace($result.Message)) { $result.Message = 'Connected.' }
    }
    catch {
        $result.Success = $false
        $result.Message = $_.Exception.Message
    }
    finally {
        Restore-PMArcGISCertificatePolicy
    }

    return $result
}
