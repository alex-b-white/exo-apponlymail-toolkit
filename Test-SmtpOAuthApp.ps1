#Requires -Version 5.1
<#
    .SYNOPSIS
    Tests app-only SMTP OAuth2 authentication and sending for Exchange Online
    using the SMTP.SendAsApp client credentials flow.

    .DESCRIPTION
    Validates the complete SMTP.SendAsApp flow as documented at:
    https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth

    Tests performed:

      1. Prerequisites
         - Org-level SMTP AUTH not globally disabled
         - Per-mailbox SMTP AUTH enabled
         - EXO service principal registered
         - FullAccess and SendAs permissions granted per mailbox

      2. Token acquisition
         - Client credentials flow
         - Scope: https://outlook.office365.com/.default
         - Token audience verified as https://outlook.office365.com
         - SMTP.SendAsApp role claim verified in token

      3. SMTP conversation per mailbox
         - TCP connect to smtp.office365.com:587
         - STARTTLS negotiation
         - EHLO — XOAUTH2 advertisement verified
         - AUTH XOAUTH2 with user= set to mailbox address
         - 235 authentication success verified
         - Optional: full message send with 250 End-of-DATA verification

    .PARAMETER ConfigFile
    Path to the JSON config file produced by New-SmtpOAuthApp.ps1.

    .PARAMETER TenantId
    Tenant ID. Required if -ConfigFile is not specified.

    .PARAMETER AppId
    Application (client) ID. Required if -ConfigFile is not specified.

    .PARAMETER ClientSecret
    Client secret. Required if -ConfigFile is not specified and not
    using a certificate.

    .PARAMETER CertThumbprint
    Certificate thumbprint. Required if -ConfigFile is not specified
    and using a certificate.

    .PARAMETER Mailboxes
    Mailbox addresses to test. Uses config file list if not specified.

    .PARAMETER SkipSmtpSend
    Perform token and auth validation only — do not send a test message.

    .PARAMETER TestRecipient
    Recipient for test messages. Defaults to the sending mailbox (loop-back).

    .PARAMETER TimeoutMs
    TCP timeout in milliseconds. Defaults to 30000.

    .EXAMPLE
    .\Test-SmtpOAuthApp.ps1 -ConfigFile ".\SmtpMailer-HR-smtp-config.json"

    .EXAMPLE
    .\Test-SmtpOAuthApp.ps1 `
        -ConfigFile    ".\SmtpMailer-HR-smtp-config.json" `
        -TestRecipient "admin@contoso.com"

    .EXAMPLE
    .\Test-SmtpOAuthApp.ps1 `
        -ConfigFile   ".\SmtpMailer-HR-smtp-config.json" `
        -SkipSmtpSend

    .NOTES
    Requires:
      - ExchangeOnlineManagement module
      - Network access to smtp.office365.com:587
      - New-SmtpOAuthApp.ps1 completed successfully
      - Windows PowerShell 5.1 or PowerShell 7+
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$ConfigFile,

    [Parameter()]
    [string]$TenantId,

    [Parameter()]
    [string]$AppId,

    [Parameter()]
    [string]$ClientSecret,

    [Parameter()]
    [string]$CertThumbprint,

    [Parameter()]
    [string[]]$Mailboxes,

    [Parameter()]
    [switch]$SkipSmtpSend,

    [Parameter()]
    [string]$TestRecipient,

    [Parameter()]
    [ValidateRange(5000, 120000)]
    [int]$TimeoutMs = 30000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:ExoResourceUrl = "https://outlook.office365.com"
$script:SmtpHost       = "smtp.office365.com"
$script:SmtpPort       = 587

# ── Compatible module versions ───────────────────────────────────────────────
# ExchangeOnlineManagement 3.8+ and Microsoft.Graph 2.25+ cannot coexist in
# the same PowerShell session due to a WAM/MSAL broker DLL conflict.
# These versions are the last known-good combination where both load cleanly.
$script:RequiredExoVersion   = "3.7.0"
$script:RequiredGraphVersion = "2.24.0"

$script:ExoResourceUrl = "https://outlook.office365.com"
$script:SmtpHost       = "smtp.office365.com"
$script:SmtpPort       = 587

#region ── Helpers ─────────────────────────────────────────────────────────────

$script:TestsPassed  = [System.Collections.Generic.List[string]]::new()
$script:TestsFailed  = [System.Collections.Generic.List[string]]::new()
$script:TestsWarned  = [System.Collections.Generic.List[string]]::new()
$script:TestsSkipped = [System.Collections.Generic.List[string]]::new()
$script:accessToken  = $null

function Assert-NotNull {
    param([object]$Value, [string]$Label)
    if ($null -eq $Value -or
        ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) {
        throw "$Label was null or empty."
    }
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("─" * 70) -ForegroundColor DarkGray
    Write-Host "  ▶  $Title" -ForegroundColor Cyan
    Write-Host ("─" * 70) -ForegroundColor DarkGray
}

function Write-TestResult {
    param(
        [string]$TestName,
        [ValidateSet("PASS","FAIL","WARN","SKIP","INFO")]
        [string]$Result,
        [string]$Detail = ""
    )
    $icon  = switch ($Result) {
        "PASS" { "✔" }; "FAIL" { "✘" }; "WARN" { "⚠" }
        "SKIP" { "○" }; "INFO" { "ℹ" }
    }
    $color = switch ($Result) {
        "PASS" { "Green"    }; "FAIL" { "Red"      }
        "WARN" { "Yellow"   }; "SKIP" { "DarkGray" }
        "INFO" { "DarkCyan" }
    }
    $line = "  $icon  [$Result] $TestName"
    if ($Detail) { $line += " — $Detail" }
    Write-Host $line -ForegroundColor $color

    switch ($Result) {
        "PASS" { $script:TestsPassed.Add($TestName)  }
        "FAIL" { $script:TestsFailed.Add($TestName)  }
        "WARN" { $script:TestsWarned.Add($TestName)  }
        "SKIP" { $script:TestsSkipped.Add($TestName) }
    }
}

function Invoke-TestBlock {
    param(
        [string]$Name,
        [scriptblock]$Test,
        [string]$ExpectedFailure
    )
    try {
        $detail = & $Test
        if ($ExpectedFailure) {
            Write-TestResult -TestName $Name -Result "FAIL" `
                -Detail "Expected failure but request succeeded."
        }
        else {
            Write-TestResult -TestName $Name -Result "PASS" `
                -Detail ($detail -as [string])
        }
    }
    catch {
        $msg = $_.Exception.Message
        if ($ExpectedFailure -and $msg -match $ExpectedFailure) {
            Write-TestResult -TestName $Name -Result "PASS" `
                -Detail "Correctly denied (matched: $ExpectedFailure)."
        }
        else {
            Write-TestResult -TestName $Name -Result "FAIL" -Detail $msg
        }
    }
}

function Decode-JwtPayload {
    param([string]$Token)
    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { throw "Token does not appear to be a JWT." }
    $padded = $parts[1].PadRight(
        $parts[1].Length + (4 - $parts[1].Length % 4) % 4, '=')
    return [Text.Encoding]::UTF8.GetString(
        [Convert]::FromBase64String($padded)) | ConvertFrom-Json
}

function Get-ValueOrDefault {
    param([object]$Value, [object]$Default)
    if ($null -ne $Value -and
        -not ($Value -is [string] -and [string]::IsNullOrEmpty($Value))) {
        return $Value
    }
    return $Default
}

#endregion

#region ── SMTP helpers ────────────────────────────────────────────────────────

function Build-XOAuth2String {
    param([string]$Username, [string]$AccessToken)
    # SOH (Start of Heading, 0x01) is the required separator in XOAUTH2.
    # Use [char]0x01 — compatible with both Windows PowerShell 5.1 and PS 7+.
    # Do NOT use `u{0001} — that syntax requires PowerShell 6.0+.
    $soh = [char]0x01
    $raw = "user=$Username${soh}auth=Bearer $AccessToken${soh}${soh}"
    return [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($raw))
}

function Invoke-SmtpConversation {
    param(
        [string]$FromAddress,
        [string]$ToAddress,
        [string]$AccessToken,
        [int]   $TimeoutMs = 30000,
        [switch]$AuthOnly
    )

    $result = [ordered]@{
        FromAddress  = $FromAddress
        ToAddress    = $ToAddress
        Success      = $false
        FailureStage = $null
        FailureCode  = $null
        Log          = [System.Collections.Generic.List[string]]::new()
    }

    function Log  { param([string]$m) $result.Log.Add($m) }
    function Fail { param([string]$stage, [string]$code)
        $result.FailureStage = $stage
        $result.FailureCode  = $code
    }

    function Read-SmtpResponse {
        param([System.IO.StreamReader]$Reader)
        $buf = [System.Text.StringBuilder]::new()
        do {
            $line = $Reader.ReadLine()
            [void]$buf.AppendLine($line)
        } while ($line -match '^\d{3}-')
        return $buf.ToString().Trim()
    }

    function Assert-Code {
        param([string]$Response, [string]$Expected, [string]$Stage)
        $code = $Response.Substring(0, [Math]::Min(3, $Response.Length))
        Log "[$Stage] S: $Response"
        if ($code -ne $Expected) {
            # Provide specific guidance for known transient store errors
            if ($code -eq "430" -and $Response -match "STOREDRV") {
                Fail -stage $Stage -code "430-STOREDRV"
            }
            else {
                Fail -stage $Stage -code $code
            }
            throw "Expected $Expected at '$Stage', got: $Response"
        }
    }

    $tcp = $ssl = $reader = $writer = $null

    try {
        Log "Connecting to $($script:SmtpHost):$($script:SmtpPort)..."
        $tcp = [Net.Sockets.TcpClient]::new()
        $tcp.ReceiveTimeout = $TimeoutMs
        $tcp.SendTimeout    = $TimeoutMs
        $tcp.Connect($script:SmtpHost, $script:SmtpPort)

        $stream = $tcp.GetStream()
        $reader = [IO.StreamReader]::new($stream)
        $writer = [IO.StreamWriter]::new($stream)
        $writer.AutoFlush = $true

        $banner = Read-SmtpResponse $reader
        Assert-Code -Response $banner -Expected "220" -Stage "Banner"

        Log "[EHLO] C: EHLO test.local"
        $writer.WriteLine("EHLO test.local")
        $ehlo1 = Read-SmtpResponse $reader
        Assert-Code -Response $ehlo1 -Expected "250" -Stage "EHLO-plain"

        if ($ehlo1 -notmatch "STARTTLS") {
            Fail -stage "STARTTLS-advertised" -code "N/A"
            throw "Server did not advertise STARTTLS."
        }
        Log "[STARTTLS-check] STARTTLS advertised."

        Log "[STARTTLS] C: STARTTLS"
        $writer.WriteLine("STARTTLS")
        $tls = Read-SmtpResponse $reader
        Assert-Code -Response $tls -Expected "220" -Stage "STARTTLS"

        Log "[TLS] Upgrading to TLS..."
        $ssl = [Net.Security.SslStream]::new($stream, $false,
            { param($s,$c,$ch,$e) $true })
        $ssl.AuthenticateAsClient($script:SmtpHost)
        Log "[TLS] Protocol: $($ssl.SslProtocol)  Cipher: $($ssl.CipherAlgorithm)"

        $reader = [IO.StreamReader]::new($ssl)
        $writer = [IO.StreamWriter]::new($ssl)
        $writer.AutoFlush = $true

        Log "[EHLO-TLS] C: EHLO test.local"
        $writer.WriteLine("EHLO test.local")
        $ehlo2 = Read-SmtpResponse $reader
        Assert-Code -Response $ehlo2 -Expected "250" -Stage "EHLO-TLS"

        if ($ehlo2 -notmatch "XOAUTH2") {
            Fail -stage "XOAUTH2-advertised" -code "N/A"
            throw ("Server did not advertise AUTH XOAUTH2 after STARTTLS. " +
                   "Ensure SmtpClientAuthenticationDisabled = `$false on " +
                   "the mailbox and the org-level setting is not globally off.")
        }
        Log "[XOAUTH2-check] AUTH XOAUTH2 advertised."

        $xoauth2 = Build-XOAuth2String `
            -Username    $FromAddress `
            -AccessToken $AccessToken
        Log "[AUTH] C: AUTH XOAUTH2 <base64 — user=$FromAddress>"
        $writer.WriteLine("AUTH XOAUTH2 $xoauth2")
        $auth = Read-SmtpResponse $reader
        Assert-Code -Response $auth -Expected "235" -Stage "AUTH"

        if ($AuthOnly) {
            $result.Success = $true
            Log "Auth-only mode — skipping message send."
            Log "[QUIT] C: QUIT"
            $writer.WriteLine("QUIT")
            $quit = Read-SmtpResponse $reader
            Log "[QUIT] S: $quit"
            return $result
        }

        Log "[MAIL FROM] C: MAIL FROM:<$FromAddress>"
        $writer.WriteLine("MAIL FROM:<$FromAddress>")
        $mf = Read-SmtpResponse $reader
        Assert-Code -Response $mf -Expected "250" -Stage "MAIL FROM"

        Log "[RCPT TO] C: RCPT TO:<$ToAddress>"
        $writer.WriteLine("RCPT TO:<$ToAddress>")
        $rt = Read-SmtpResponse $reader
        Assert-Code -Response $rt -Expected "250" -Stage "RCPT TO"

        Log "[DATA] C: DATA"
        $writer.WriteLine("DATA")
        $data = Read-SmtpResponse $reader
        Assert-Code -Response $data -Expected "354" -Stage "DATA"

        $ts    = Get-Date -Format "u"
        $msgId = [Guid]::NewGuid().ToString()
        $writer.WriteLine("From: <$FromAddress>")
        $writer.WriteLine("To: <$ToAddress>")
        $writer.WriteLine("Subject: [Test-SmtpOAuthApp] SMTP.SendAsApp test — $ts")
        $writer.WriteLine("Message-ID: <$msgId@test-smtpoauthapp>")
        $writer.WriteLine("Date: $(Get-Date -Format 'ddd, dd MMM yyyy HH:mm:ss zzz')")
        $writer.WriteLine("Content-Type: text/plain; charset=utf-8")
        $writer.WriteLine("")
        $writer.WriteLine("Automated SMTP.SendAsApp OAuth2 test.")
        $writer.WriteLine("Sender   : $FromAddress")
        $writer.WriteLine("Recipient: $ToAddress")
        $writer.WriteLine("Time     : $ts")
        $writer.WriteLine(".")
        $dot = Read-SmtpResponse $reader
        Assert-Code -Response $dot -Expected "250" -Stage "End-of-DATA"

        Log "[QUIT] C: QUIT"
        $writer.WriteLine("QUIT")
        $quit = Read-SmtpResponse $reader
        Log "[QUIT] S: $quit"

        $result.Success = $true
        Log "Message accepted."
    }
    catch {
        if (-not $result.FailureStage) {
            $result.FailureStage = "Unknown"
            $result.FailureCode  = "Exception"
        }
        $result.Log.Add("Exception: $_")
    }
    finally {
        if ($writer) { try { $writer.Dispose() } catch {} }
        if ($reader) { try { $reader.Dispose() } catch {} }
        if ($ssl)    { try { $ssl.Dispose()    } catch {} }
        if ($tcp)    { try { $tcp.Close()      } catch {} }
    }

    return $result
}

#endregion

#region ── Banner ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("═" * 70) -ForegroundColor DarkCyan
Write-Host ("  Test-SmtpOAuthApp  │  " +
            "App-Only SMTP OAuth2 Validation (SMTP.SendAsApp)") `
    -ForegroundColor Cyan
Write-Host ("═" * 70) -ForegroundColor DarkCyan
Write-Host "  Config File : $(if ($ConfigFile) { $ConfigFile } else { '(not specified)' })"
Write-Host "  SMTP Send   : $(if ($SkipSmtpSend) { 'Skipped' } else { 'Enabled' })"
Write-Host ("═" * 70) -ForegroundColor DarkCyan

#endregion

#region ── Load Config ─────────────────────────────────────────────────────────

Write-Section "Loading configuration"

$cfg = $null

if ($ConfigFile) {
    Invoke-TestBlock -Name "Config file exists" -Test {
        if (-not (Test-Path $ConfigFile)) {
            throw "File not found: $ConfigFile"
        }
        "Found: $ConfigFile"
    }

    Invoke-TestBlock -Name "Config file parse" -Test {
        $script:cfg = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        "Parsed successfully."
    }

    $cfg = $script:cfg

    if ($cfg) {
        if (-not $TenantId      -and $cfg.TenantId)      { $TenantId      = $cfg.TenantId      }
        if (-not $AppId         -and $cfg.AppId)          { $AppId         = $cfg.AppId          }
        if (-not $ClientSecret  -and $cfg.ClientSecret)   { $ClientSecret  = $cfg.ClientSecret   }
        if (-not $CertThumbprint -and $cfg.CertThumbprint){ $CertThumbprint = $cfg.CertThumbprint }
        if (-not $Mailboxes     -and $cfg.Mailboxes)      { $Mailboxes     = $cfg.Mailboxes      }
        if ($cfg.SmtpHost)      { $script:SmtpHost      = $cfg.SmtpHost      }
        if ($cfg.SmtpPort)      { $script:SmtpPort      = $cfg.SmtpPort      }
        if ($cfg.TokenResource) { $script:ExoResourceUrl = $cfg.TokenResource }

        Write-TestResult -TestName "Config values" -Result "INFO" `
            -Detail ("TenantId=$TenantId  AppId=$AppId  " +
                     "Mailboxes=$($Mailboxes -join ',')")
    }
}

if (-not $TenantId -or -not $AppId -or -not $Mailboxes) {
    Write-TestResult -TestName "Required parameters" -Result "FAIL" `
        -Detail ("TenantId, AppId, and Mailboxes are all required. " +
                 "Provide -ConfigFile or specify them directly.")
    exit 1
}

if (-not $ClientSecret -and -not $CertThumbprint) {
    Write-TestResult -TestName "Credential" -Result "FAIL" `
        -Detail "Either -ClientSecret or -CertThumbprint must be provided."
    exit 1
}

#endregion

#region ── Connect to Exchange Online ──────────────────────────────────────────

Write-Section "Connecting to Exchange Online"

try {
    $null = Get-OrganizationConfig -ErrorAction Stop
    Write-TestResult -TestName "Exchange Online session" -Result "PASS" `
        -Detail "Already connected."
}
catch {
    try {
        Import-Module ExchangeOnlineManagement `
            -RequiredVersion $script:RequiredExoVersion `
            -ErrorAction Stop

        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-TestResult -TestName "Exchange Online session" -Result "PASS" `
            -Detail "Connected."
    }
    catch {
        Write-TestResult -TestName "Exchange Online session" -Result "WARN" `
            -Detail ("Could not connect — EXO checks will be skipped. " +
                     "Error: $_")
    }
}

#endregion

#region ── Prerequisite Checks ────────────────────────────────────────────────

Write-Section "Organisation-level SMTP AUTH"

Invoke-TestBlock -Name "Org SMTP AUTH setting (informational)" -Test {
    $transport = Get-TransportConfig -ErrorAction Stop
    $val       = $transport.SmtpClientAuthenticationDisabled
    if ($val -eq $true) {
        Write-TestResult -TestName "Org SMTP AUTH globally disabled" -Result "INFO" `
            -Detail ("Org-level SMTP AUTH is disabled, but per-mailbox override " +
                     "(SmtpClientAuthenticationDisabled = `$false) takes precedence. " +
                     "No action required.")
    }
    "SmtpClientAuthenticationDisabled = $val (per-mailbox settings apply regardless)."
}


Write-Section "Per-mailbox SMTP AUTH settings"

foreach ($addr in $Mailboxes) {
    Invoke-TestBlock -Name "SMTP AUTH enabled: $addr" -Test {
        $cas = Get-CASMailbox -Identity $addr -ErrorAction Stop
        $val = $cas.SmtpClientAuthenticationDisabled
        if ($val -eq $true) {
            throw ("SmtpClientAuthenticationDisabled = True for '$addr'. " +
                   "Fix: Set-CASMailbox '$addr' " +
                   "-SmtpClientAuthenticationDisabled `$false")
        }
        "SmtpClientAuthenticationDisabled = $val."
    }
}

Write-Section "Exchange Online service principal registration"

Invoke-TestBlock -Name "EXO service principal registered" -Test {
    $exoSp = Get-ServicePrincipal -ErrorAction SilentlyContinue |
        Where-Object { $_.AppId -eq $AppId } |
        Select-Object -First 1

    if (-not $exoSp) {
        throw ("EXO service principal not found for AppId '$AppId'. " +
               "Run New-SmtpOAuthApp.ps1 to register it.")
    }
    "Identity: $($exoSp.Identity)"
}

Write-Section "Mailbox permissions"

foreach ($addr in $Mailboxes) {
    Invoke-TestBlock -Name "FullAccess permission: $addr" -Test {
        $exoSp = Get-ServicePrincipal -ErrorAction Stop |
            Where-Object { $_.AppId -eq $AppId } |
            Select-Object -First 1

        if (-not $exoSp) { throw "EXO service principal not found." }

        # Get-MailboxPermission returns the User/trustee in various formats:
        #   - DOMAIN\<guid>
        #   - <guid>
        #   - <displayname>
        # Match on the GUID portion of the Identity which is always present.
        $spGuid = $exoSp.Identity   # already a GUID string

        $perms = Get-MailboxPermission -Identity $addr -ErrorAction Stop |
            Where-Object {
                ($_.User -like "*$spGuid*" -or
                 $_.User -like "*$($exoSp.DisplayName)*") -and
                $_.AccessRights -like "*FullAccess*" -and
                $_.Deny -ne $true
            }

        if (-not $perms) {
            # Show what IS there, to aid diagnosis
            $existing = (Get-MailboxPermission -Identity $addr |
                Where-Object { $_.AccessRights -like "*FullAccess*" } |
                Select-Object -ExpandProperty User) -join "; "
            throw ("FullAccess not granted for '$addr'. " +
                   "Current FullAccess trustees: [$existing]. " +
                   "Run: Add-MailboxPermission -Identity '$addr' " +
                   "-User '$spGuid' -AccessRights FullAccess")
        }
        "FullAccess confirmed (trustee: $($perms[0].User))."
    }

    Invoke-TestBlock -Name "SendAs permission: $addr" -Test {
        $exoSp = Get-ServicePrincipal -ErrorAction Stop |
            Where-Object { $_.AppId -eq $AppId } |
            Select-Object -First 1

        if (-not $exoSp) { throw "EXO service principal not found." }

        $perms = Get-RecipientPermission -Identity $addr -ErrorAction Stop |
            Where-Object { $_.Trustee -like "*$($exoSp.Identity)*" -and
                           $_.AccessRights -like "*SendAs*" }

        if (-not $perms) {
            throw ("SendAs not granted for '$addr'. " +
                   "Run: Add-RecipientPermission -Identity '$addr' " +
                   "-Trustee '$($exoSp.Identity)' -AccessRights SendAs")
        }
        "SendAs confirmed."
    }
}

#endregion

#region ── Token Acquisition ──────────────────────────────────────────────────

Write-Section "Token acquisition (client credentials)"

Invoke-TestBlock -Name "Acquire access token" -Test {
    if ($CertThumbprint) {
        $cert = Get-Item "Cert:\CurrentUser\My\$CertThumbprint" -ErrorAction Stop

        $now     = [DateTimeOffset]::UtcNow
        $header  = [Convert]::ToBase64String(
            [Text.Encoding]::UTF8.GetBytes(
                '{"alg":"RS256","typ":"JWT","x5t":"' +
                [Convert]::ToBase64String($cert.GetCertHash()) +
                '"}'
            )
        ).TrimEnd('=').Replace('+','-').Replace('/','_')

        $payload = [Convert]::ToBase64String(
            [Text.Encoding]::UTF8.GetBytes(
                (ConvertTo-Json @{
                    aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
                    exp = $now.AddMinutes(5).ToUnixTimeSeconds()
                    iss = $AppId
                    jti = [Guid]::NewGuid().ToString()
                    nbf = $now.ToUnixTimeSeconds()
                    sub = $AppId
                } -Compress)
            )
        ).TrimEnd('=').Replace('+','-').Replace('/','_')

        $toSign   = "$header.$payload"
        $rsa      = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        $sigBytes = $rsa.SignData(
            [Text.Encoding]::UTF8.GetBytes($toSign),
            [Security.Cryptography.HashAlgorithmName]::SHA256,
            [Security.Cryptography.RSASignaturePadding]::Pkcs1
        )
        $sig = [Convert]::ToBase64String($sigBytes).TrimEnd('=').Replace('+','-').Replace('/','_')

        $tokenBody = @{
            client_id             = $AppId
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            client_assertion      = "$toSign.$sig"
            scope                 = "$script:ExoResourceUrl/.default"
            grant_type            = "client_credentials"
        }
    }
    else {
        $tokenBody = @{
            client_id     = $AppId
            client_secret = $ClientSecret
            scope         = "$script:ExoResourceUrl/.default"
            grant_type    = "client_credentials"
        }
    }

    $tokenResponse = Invoke-RestMethod `
        -Method Post `
        -Uri    "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Body   $tokenBody `
        -ErrorAction Stop

    Assert-NotNull -Value $tokenResponse.access_token -Label "Access token"
    $script:accessToken = $tokenResponse.access_token

    $expiry = (Get-Date).AddSeconds($tokenResponse.expires_in)
    "Token acquired. Expires: $($expiry.ToString('HH:mm:ss'))"
}

#endregion

#region ── Token Claim Verification ───────────────────────────────────────────

Write-Section "Token claim verification"

Invoke-TestBlock -Name "Token audience" -Test {
    if (-not $script:accessToken) { throw "No token available." }
    $payload = Decode-JwtPayload -Token $script:accessToken

    if ($payload.aud -ne $script:ExoResourceUrl) {
        throw ("aud='$($payload.aud)' — expected '$script:ExoResourceUrl'. " +
               "Ensure scope is '$script:ExoResourceUrl/.default', " +
               "not 'https://graph.microsoft.com/.default'.")
    }
    "aud=$($payload.aud)"
}

Invoke-TestBlock -Name "Token contains SMTP.SendAsApp role" -Test {
    if (-not $script:accessToken) { throw "No token available." }
    $payload = Decode-JwtPayload -Token $script:accessToken

    # App-only tokens carry permissions in the 'roles' claim, not 'scp'
    $roles = ""
    if ($null -ne $payload.roles) {
        if ($payload.roles -is [array]) {
            $roles = $payload.roles -join " "
        }
        else {
            $roles = $payload.roles
        }
    }

    if ($roles -notmatch "SMTP\.SendAsApp") {
        throw ("SMTP.SendAsApp not found in roles claim. " +
               "roles='$roles'. " +
               "Ensure the SMTP.SendAsApp application permission has been " +
               "added to the app registration on the Office 365 Exchange " +
               "Online resource (not Microsoft Graph) and admin consent " +
               "has been granted.")
    }
    "roles=$roles"
}

Invoke-TestBlock -Name "Token expiry" -Test {
    if (-not $script:accessToken) { throw "No token available." }
    $payload = Decode-JwtPayload -Token $script:accessToken
    $expiry  = (Get-Date "1970-01-01T00:00:00Z").AddSeconds($payload.exp)
    if ((Get-Date) -gt $expiry) {
        throw "Token has already expired: $($expiry.ToLocalTime())"
    }
    "Expires: $($expiry.ToLocalTime().ToString('HH:mm:ss'))"
}

#endregion

#region ── Per-Mailbox SMTP Tests ─────────────────────────────────────────────

foreach ($addr in $Mailboxes) {

    Write-Section "SMTP OAuth2: $addr"

    if (-not $script:accessToken) {
        Write-TestResult -TestName "SMTP auth: $addr" -Result "SKIP" `
            -Detail "Skipped — no access token."
        if (-not $SkipSmtpSend) {
            Write-TestResult -TestName "SMTP send: $addr" -Result "SKIP" `
                -Detail "Skipped — no access token."
        }
        continue
    }

    # ── Auth-only test ─────────────────────────────────────────────────────
    Invoke-TestBlock -Name "SMTP auth: $addr" -Test {
        $smtpResult = Invoke-SmtpConversation `
            -FromAddress $addr `
            -ToAddress   $addr `
            -AccessToken $script:accessToken `
            -TimeoutMs   $TimeoutMs `
            -AuthOnly

        $smtpResult.Log | ForEach-Object {
            Write-Host "    $_" -ForegroundColor DarkGray
        }

        if (-not $smtpResult.Success) {
            throw ("SMTP auth failed at stage '$($smtpResult.FailureStage)' " +
                   "code '$($smtpResult.FailureCode)'.")
        }

        "235 Authentication successful."
    }

    # ── Send test ──────────────────────────────────────────────────────────
    if ($SkipSmtpSend) {
        Write-TestResult -TestName "SMTP send: $addr" -Result "SKIP" `
            -Detail "Skipped via -SkipSmtpSend."
    }
    else {
        $recipient = if ($TestRecipient) { $TestRecipient } else { $addr }

        Invoke-TestBlock -Name "SMTP send: $addr → $recipient" -Test {
            $smtpResult = Invoke-SmtpConversation `
                -FromAddress $addr `
                -ToAddress   $recipient `
                -AccessToken $script:accessToken `
                -TimeoutMs   $TimeoutMs

            $smtpResult.Log | ForEach-Object {
                Write-Host "    $_" -ForegroundColor DarkGray
            }

            if (-not $smtpResult.Success) {
                throw ("SMTP send failed at stage " +
                       "'$($smtpResult.FailureStage)' " +
                       "code '$($smtpResult.FailureCode)'.")
            }

            "Message accepted at End-of-DATA."
        }
    }
}

#endregion

#region ── Final Results ───────────────────────────────────────────────────────

Write-Host ""
Write-Host ("═" * 70) -ForegroundColor DarkGray
Write-Host "  TEST RESULTS" -ForegroundColor White
Write-Host ("═" * 70) -ForegroundColor DarkGray

Write-Host ("  ✔  Passed  : {0,3}" -f $script:TestsPassed.Count)  -ForegroundColor Green
Write-Host ("  ✘  Failed  : {0,3}" -f $script:TestsFailed.Count)  -ForegroundColor Red
Write-Host ("  ⚠  Warned  : {0,3}" -f $script:TestsWarned.Count)  -ForegroundColor Yellow
Write-Host ("  ○  Skipped : {0,3}" -f $script:TestsSkipped.Count) -ForegroundColor DarkGray

if ($script:TestsFailed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Failed tests:" -ForegroundColor Red
    $script:TestsFailed | ForEach-Object {
        Write-Host "    • $_" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "  Troubleshooting guide:" -ForegroundColor Yellow
    $failedNames = $script:TestsFailed -join " "

    if ($failedNames -match "Token audience") {
        Write-Host ("    Token:  Scope must be " +
                    "'https://outlook.office365.com/.default'") `
            -ForegroundColor DarkYellow
        Write-Host ("    Token:  Do NOT use 'https://graph.microsoft.com/.default'") `
            -ForegroundColor DarkYellow
    }
    if ($failedNames -match "SMTP.SendAsApp") {
        Write-Host ("    Perm:   Add SMTP.SendAsApp APPLICATION permission on") `
            -ForegroundColor DarkYellow
        Write-Host ("    Perm:   'Office 365 Exchange Online' resource " +
                    "(AppId 00000002-0000-0ff1-ce00-000000000000)") `
            -ForegroundColor DarkYellow
        Write-Host ("    Perm:   NOT on Microsoft Graph") `
            -ForegroundColor DarkYellow
        Write-Host ("    Perm:   Grant admin consent after adding the permission") `
            -ForegroundColor DarkYellow
    }
    if ($failedNames -match "EXO service principal") {
        Write-Host ("    EXO SP: Run New-ServicePrincipal in Exchange Online") `
            -ForegroundColor DarkYellow
        Write-Host ("    EXO SP: Use the Enterprise App Object ID " +
                    "(service principal), not the App Registration Object ID") `
            -ForegroundColor DarkYellow
    }
    if ($failedNames -match "FullAccess|SendAs") {
        Write-Host ("    Perms:  Run Add-MailboxPermission and " +
                    "Add-RecipientPermission for each mailbox") `
            -ForegroundColor DarkYellow
    }
    if ($failedNames -match "SMTP auth|SMTP send") {
        Write-Host ("    SMTP:   535 at AUTH — check token claims and " +
                    "EXO service principal registration") `
            -ForegroundColor DarkYellow
        Write-Host ("    SMTP:   XOAUTH2 not advertised — check per-mailbox " +
                    "and org-level SMTP AUTH settings") `
            -ForegroundColor DarkYellow
    }
    if ($failedNames -match "SMTP send") {
    # Add this alongside the existing 535 guidance:
    Write-Host ("    SMTP:   430 STOREDRV MapiExceptionLogonFailed — " +
                "Wait 15-20 min for MAPI store propagation") `
        -ForegroundColor DarkYellow
    }
}

Write-Host ""
Write-Host ("═" * 70) -ForegroundColor DarkGray

if ($script:TestsFailed.Count -eq 0 -and $script:TestsWarned.Count -eq 0) {
    Write-Host "  ✔  ALL TESTS PASSED" -ForegroundColor Green
}
elseif ($script:TestsFailed.Count -eq 0) {
    Write-Host "  ⚠  ALL TESTS PASSED WITH WARNINGS" -ForegroundColor Yellow
}
else {
    Write-Host "  ✘  SOME TESTS FAILED" -ForegroundColor Red
}

Write-Host ("═" * 70) -ForegroundColor DarkGray

exit $(if ($script:TestsFailed.Count -gt 0) { 1 } else { 0 })

#endregion
