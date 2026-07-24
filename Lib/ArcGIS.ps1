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
#    - POST for generateToken, data/findItems, data/validateDataItem and
#      logs/query ONLY. All four are POST because the Admin API requires
#      it for these operations (a token, or a JSON body too large for a
#      query string, must travel in the request), NOT because they change
#      anything: generateToken issues a short-lived token, data/findItems
#      only lists what is already registered, data/validateDataItem only
#      retests a connection the site already has stored credentials for
#      using those credentials, and logs/query only reads the log the
#      site already keeps - see Checks\A3-ArcGISData.ps1 (AGSDATA) and
#      Checks\A4-ArcGISLog.ps1 (AGSLOG). AGSDATA is also careful about
#      what it never keeps: findItems returns each item's connection
#      string, encrypted password included, and that never leaves memory.
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

# Same job as Get-PMArcGISRoot, for the Portal for ArcGIS site that a
# federated ArcGIS Server trusts for sign-in. The forms people paste:
#
#   https://portal.example.go.th/portal         -> unchanged
#   https://portal.example.go.th                -> .../portal (default site name)
#   https://portal.example.go.th/portal/home    -> .../portal (trailing app page tolerated)
#   https://portal.example.go.th/portal/sharing/rest -> .../portal (already-typed API path tolerated)
function Get-PMArcGISPortalRoot {
    param([Parameter(Mandatory)][string]$Url)

    $u = $Url.Trim()
    if ([string]::IsNullOrWhiteSpace($u)) { throw 'The Portal for ArcGIS URL is empty.' }

    if ($u -notmatch '^https?://') { $u = 'https://' + $u }
    $u = $u.TrimEnd('/')
    $u = $u -replace '(?i)/sharing(/rest)?/?$', ''
    $u = $u -replace '(?i)/home/?$', ''
    $u = $u.TrimEnd('/')

    $uri = $null
    if (-not [Uri]::TryCreate($u, [UriKind]::Absolute, [ref]$uri)) {
        throw "Not a valid URL: $Url"
    }

    # Host with no path at all: the default Portal site name is the only
    # sensible guess, and it is right on a default install.
    if ([string]::IsNullOrWhiteSpace($uri.AbsolutePath) -or $uri.AbsolutePath -eq '/') {
        $u = $u + '/portal'
    }

    return $u
}

function Get-PMArcGISPortalTokenUrl {
    param([Parameter(Mandatory)][string]$PortalRoot)
    return ($PortalRoot + '/sharing/rest/generateToken')
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
        [ValidateSet('Server', 'Portal')][string]$AuthMode = 'Server',
        [string]$PortalUrl = '',
        [string]$ConfigDir
    )
    $path = Get-PMArcGISConnectionPath -ConfigDir $ConfigDir
    $dir  = Split-Path -Parent $path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $resolvedPortalUrl = ''
    if ($AuthMode -eq 'Portal') { $resolvedPortalUrl = Get-PMArcGISPortalRoot -Url $PortalUrl }

    [pscustomobject]@{
        Url       = Get-PMArcGISRoot -Url $Url
        AuthMode  = $AuthMode
        PortalUrl = $resolvedPortalUrl
        Username  = $Username
        Password  = ConvertFrom-SecureString -SecureString $Password
        SavedBy   = "$env:USERDOMAIN\$env:USERNAME"
        SavedOn   = $env:COMPUTERNAME
        SavedAt   = (Get-Date).ToString('s')
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

    # AuthMode/PortalUrl were added after this file format shipped - a
    # connection saved by an older PMtools has neither property, and reads
    # back as plain server-tier auth (today's only behavior back then).
    $authMode = 'Server'
    if ($saved.PSObject.Properties['AuthMode'] -and -not [string]::IsNullOrWhiteSpace([string]$saved.AuthMode)) {
        $authMode = [string]$saved.AuthMode
    }
    $portalUrl = ''
    if ($saved.PSObject.Properties['PortalUrl']) { $portalUrl = [string]$saved.PortalUrl }

    return [pscustomobject]@{
        Url       = [string]$saved.Url
        AuthMode  = $authMode
        PortalUrl = $portalUrl
        Username  = [string]$saved.Username
        Password  = $secure
        SavedBy   = [string]$saved.SavedBy
        SavedOn   = [string]$saved.SavedOn
        SavedAt   = [string]$saved.SavedAt
        Path      = $path
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
$Script:PMArcGISTokenKey    = ''

function Clear-PMArcGISToken {
    $Script:PMArcGISToken       = $null
    $Script:PMArcGISTokenExpiry = [datetime]::MinValue
    $Script:PMArcGISTokenKey    = ''
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
        [ValidateSet('Server', 'Portal')][string]$AuthMode = 'Server',
        [string]$PortalUrl = '',
        [int]$ExpirationMinutes = 60,
        [int]$TimeoutSec = 30,
        [switch]$Force
    )

    # generateToken is hit on the ArcGIS Server itself for a standalone or
    # server-tier login, or on its federated Portal instead when the saved
    # connection says so - a federated server does not accept sign-in for a
    # Portal-only named user at its own /admin/generateToken, only Portal
    # does. Either way the resulting token is then used against the SAME
    # server Admin API root everywhere else in this file; only where the
    # token itself comes from changes.
    if ($AuthMode -eq 'Portal') {
        if ([string]::IsNullOrWhiteSpace($PortalUrl)) { throw 'Portal-federated auth is selected but no Portal URL is set.' }
        $url = Get-PMArcGISPortalTokenUrl -PortalRoot (Get-PMArcGISPortalRoot -Url $PortalUrl)
    }
    else {
        $url = Get-PMArcGISAdminUrl -Root $Root -Path 'generateToken'
    }
    $tokenKey = '{0}|{1}|{2}' -f $AuthMode, $Root, $url

    if (-not $Force -and $Script:PMArcGISToken -and $Script:PMArcGISTokenKey -eq $tokenKey -and
        (Get-Date) -lt $Script:PMArcGISTokenExpiry) {
        return $Script:PMArcGISToken
    }

    Initialize-PMArcGISTransport
    if (Get-PMSetting -Path 'ArcGIS.IgnoreCertificateErrors' -Default $true) { Disable-PMArcGISCertificateCheck }

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
    $signInTarget = if ($AuthMode -eq 'Portal') { 'Portal for ArcGIS' } else { 'ArcGIS Server' }
    if ($resp.PSObject.Properties['error']) {
        $msg = $resp.error.message
        if ($resp.error.details) { $msg = $msg + ' - ' + ($resp.error.details -join '; ') }
        throw ("{0} rejected the sign-in: {1}" -f $signInTarget, $msg)
    }
    if (-not $resp.token) {
        throw ("{0} returned no token and no error. Check that the URL points at a {0} site." -f $signInTarget)
    }
    $newToken = [string]$resp.token

    # A token-shaped, error-free generateToken response is not by itself
    # proof the password was right - it only proves the Portal's token
    # service accepted the request. /admin/info never requires a token at
    # all, and the one Admin API call that does (/admin/machines, in
    # Test-PMArcGISConnection) treats ANY failure there as "signed in but
    # low-privilege" rather than "never signed in" - so a Portal that ever
    # hands back a token without really authenticating the caller would
    # read as a successful connection. Confirm the token actually
    # identifies a signed-in Portal user via community/self, which errors
    # for any token that is not genuinely tied to an account, before this
    # function will hand the token to anything else.
    if ($AuthMode -eq 'Portal') {
        $selfUrl = (Get-PMArcGISPortalRoot -Url $PortalUrl) + '/sharing/rest/community/self'
        try {
            $self = Invoke-RestMethod -Uri $selfUrl -Method Get -Body @{ f = 'json'; token = $newToken } -TimeoutSec $TimeoutSec -ErrorAction Stop
        }
        catch {
            throw ("Portal for ArcGIS issued a token but it could not be verified: {0}" -f $_.Exception.Message)
        }
        if ($self.PSObject.Properties['error']) {
            $msg = $self.error.message
            if ($self.error.details) { $msg = $msg + ' - ' + ($self.error.details -join '; ') }
            throw ("Portal for ArcGIS rejected the sign-in: {0}" -f $msg)
        }
        if (-not $self.PSObject.Properties['username'] -or [string]::IsNullOrWhiteSpace([string]$self.username)) {
            throw 'Portal for ArcGIS issued a token that does not identify a signed-in user. Check the username and password.'
        }
    }

    $Script:PMArcGISToken     = $newToken
    $Script:PMArcGISTokenKey  = $tokenKey
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
        [int]$TimeoutSec = 60,

        # data/validateDataItem answers {"status":"error","messages":[...]}
        # for a data store that genuinely cannot connect - a normal, expected
        # result its caller (AGSDATA) wants to read and report itself, not a
        # failed API call. Every other operation passes this switch, so the
        # check below stays on by default.
        [switch]$AllowStatusError
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

    # Most Admin API operations that fail answer {"status":"error",
    # "messages":[...]} in an HTTP 200 body instead of the {"error":{...}}
    # shape above - confirmed on a real federated site: /admin/info,
    # /admin/machines, data/findItems and logs/query all rejected an
    # unrecognised token this way. Left undetected, each one read as
    # ordinary data to its caller instead of a failure: /admin/info with no
    # currentversion, /admin/machines with an empty list, findItems with one
    # bogus null item (@($null) is a one-element array in PowerShell, not an
    # empty one), logs/query with zero messages - a rejected token looked
    # exactly like a healthy, empty site everywhere rather than a clear
    # error. Checked explicitly here so every caller gets the same failure
    # instead of each having to notice and work around it independently -
    # see A4-ArcGISLog.ps1's history for the one place this was already
    # worked around locally before this generic check existed.
    if (-not $AllowStatusError -and $resp -and $resp.PSObject.Properties['status'] -and [string]$resp.status -eq 'error') {
        $msg = 'unknown error'
        if ($resp.PSObject.Properties['messages'] -and $resp.messages) { $msg = (@($resp.messages) -join '; ') }
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
                               -AuthMode $conn.AuthMode -PortalUrl $conn.PortalUrl `
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
        [ValidateSet('Server', 'Portal')][string]$AuthMode = 'Server',
        [string]$PortalUrl = '',
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
                                   -AuthMode $AuthMode -PortalUrl $PortalUrl `
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
