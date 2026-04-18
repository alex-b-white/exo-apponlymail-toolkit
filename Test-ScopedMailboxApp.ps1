#Requires -Version 5.1
<#
    .SYNOPSIS
    Tests an Entra ID app registration's scoped mailbox read and send access.

    .DESCRIPTION
    This script performs the following tests:

    RBAC Tests (Exchange Online):
    - Verifies management scope exists and has the correct filter
    - Verifies role assignments exist for Mail.Read and Mail.Send
    - Verifies each mailbox has the correct custom attribute stamped
    - Verifies each mailbox is in-scope via Test-ServicePrincipalAuthorization

    API Tests (Microsoft Graph):
    - Acquires an access token using client credentials
    - Reads messages from each in-scope mailbox
    - Sends a test message from each in-scope mailbox
    - Attempts to read from an out-of-scope mailbox (expects 403/404)
    - Attempts to send from an out-of-scope mailbox (expects 403/404)

    .PARAMETER AppName
    Display name of the app registration to test.

    .PARAMETER Mailboxes
    Array of in-scope mailbox SMTP addresses to test against.

    .PARAMETER OutOfScopeMailbox
    Optional. A mailbox that EXISTS in Exchange Online but is NOT in the
    scoped group. Used to verify access is correctly denied.
    Must be a real EXO mailbox — a user with no mailbox will be skipped.

    .PARAMETER ConfigFile
    Path to the JSON config file produced by New-ScopedMailboxApp.ps1.
    If provided, AppId, TenantId and credentials are read from it automatically.

    .PARAMETER TenantId
    Entra ID tenant ID. Required if -ConfigFile is not specified.

    .PARAMETER AppId
    Application (client) ID. Required if -ConfigFile is not specified.

    .PARAMETER ClientSecret
    Client secret. Required if -ConfigFile is not specified and not using a cert.

    .PARAMETER CertThumbprint
    Certificate thumbprint. Required if -ConfigFile is not specified and
    using a cert.

    .PARAMETER SkipApiTests
    Skip the live Graph API call tests and only run RBAC verification.

    .PARAMETER TestRecipient
    Email address to send test messages to.
    Defaults to the sending mailbox itself (loop-back).

    .EXAMPLE
    .\Test-ScopedMailboxApp.ps1 `
    -AppName    "MailApp-HR" `
    -Mailboxes  @("hr@contoso.com", "payroll@contoso.com") `
    -ConfigFile ".\MailApp-HR-config.json" `
    -OutOfScopeMailbox "finance@contoso.com"

    .EXAMPLE
    .\Test-ScopedMailboxApp.ps1 `
    -AppName      "MailApp-HR" `
    -Mailboxes    @("hr@contoso.com") `
    -ConfigFile   ".\MailApp-HR-config.json" `
    -SkipApiTests

    .NOTES
    Requires:
    - ExchangeOnlineManagement module
    - An active Exchange Online session (or will prompt to connect)
    
    Out-of-scope mailbox must be a real EXO mailbox that is not stamped
    with the custom attribute used by this app's management scope.
#>

[CmdletBinding()]
param (
  [Parameter(Mandatory)]
  [ValidateNotNullOrEmpty()]
  [string]$AppName,

  [Parameter(Mandatory)]
  [ValidateNotNullOrEmpty()]
  [string[]]$Mailboxes,

  [Parameter()]
  [string]$OutOfScopeMailbox,

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
  [switch]$SkipApiTests,

  [Parameter()]
  [string]$TestRecipient
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Compatible module versions ───────────────────────────────────────────────
# ExchangeOnlineManagement 3.8+ and Microsoft.Graph 2.25+ cannot coexist in
# the same PowerShell session due to a WAM/MSAL broker DLL conflict.
# These versions are the last known-good combination where both load cleanly.
$script:RequiredExoVersion   = "3.7.0"
$script:RequiredGraphVersion = "2.24.0"

#region ── Helpers ─────────────────────────────────────────────────────────────
# All helper functions must be defined before any code that calls them.

$script:TestsPassed = [System.Collections.Generic.List[string]]::new()
$script:TestsFailed = [System.Collections.Generic.List[string]]::new()
$script:TestsWarned = [System.Collections.Generic.List[string]]::new()
$script:accessToken = $null

function Assert-NotNull {
  param(
    [object]$Value,
    [string]$Label
  )
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
    "PASS" { $script:TestsPassed.Add($TestName) }
    "FAIL" { $script:TestsFailed.Add($TestName) }
    "WARN" { $script:TestsWarned.Add($TestName) }
  }
}

function Invoke-TestBlock {
  <#
      .SYNOPSIS
      Runs a test scriptblock and records PASS/FAIL.
      If -ExpectedFailure is set, an exception whose message matches the
      pattern is treated as PASS; success is treated as FAIL.
  #>
  param(
    [string]$Name,
    [scriptblock]$Test,
    [string]$ExpectedFailure
  )
  try {
    $detail = & $Test
    if ($ExpectedFailure) {
      Write-TestResult -TestName $Name -Result "FAIL" `
      -Detail "Expected access to be denied but request succeeded."
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

#endregion

#region ── Banner ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("═" * 70) -ForegroundColor DarkCyan
Write-Host "  Test-ScopedMailboxApp  │  Validation & Live API Testing" `
-ForegroundColor Cyan
Write-Host ("═" * 70) -ForegroundColor DarkCyan
Write-Host "  App Name       : $AppName"
Write-Host "  Mailboxes      : $($Mailboxes -join ', ')"
Write-Host "  Out-of-Scope   : $(if ($OutOfScopeMailbox) { $OutOfScopeMailbox } else { '(not specified)' })"
Write-Host "  API Tests      : $(if ($SkipApiTests) { 'Skipped' } else { 'Enabled' })"
Write-Host ("═" * 70) -ForegroundColor DarkCyan

#endregion

#region ── Derived Names ───────────────────────────────────────────────────────

$ScopeName       = "$AppName-Scope"
$MESGName        = "$AppName-MESG"
$CustomAttribute = "CustomAttribute1"   # overridden from config below if present
$AttrValue       = $AppName             # overridden from config below if present

#endregion

#region ── Load Config File ────────────────────────────────────────────────────

$cfg = $null

if ($ConfigFile) {
  Write-Section "Loading configuration from file"

  if (-not (Test-Path $ConfigFile)) {
    Write-TestResult -TestName "Config file exists" -Result "FAIL" `
    -Detail "File not found: $ConfigFile"
    exit 1
  }

  try {
    $cfg = Get-Content $ConfigFile -Raw | ConvertFrom-Json
    Write-TestResult -TestName "Config file loaded" -Result "PASS" `
    -Detail $ConfigFile

    if (-not $TenantId      -and $cfg.TenantId)        { $TenantId       = $cfg.TenantId }
    if (-not $AppId         -and $cfg.AppId)            { $AppId          = $cfg.AppId }
    if (-not $ClientSecret  -and $cfg.ClientSecret)     { $ClientSecret   = $cfg.ClientSecret }
    if (-not $CertThumbprint -and $cfg.CertThumbprint)  { $CertThumbprint = $cfg.CertThumbprint }
    if ($cfg.CustomAttribute) { $CustomAttribute = $cfg.CustomAttribute }
    if ($cfg.AttributeValue)  { $AttrValue       = $cfg.AttributeValue }

    Write-TestResult -TestName "Config values" -Result "INFO" `
    -Detail ("TenantId=$TenantId  AppId=$AppId  " +
      "CredType=$($cfg.CredentialType)  " +
    "Attr=$CustomAttribute='$AttrValue'")
  }
  catch {
    Write-TestResult -TestName "Config file parse" -Result "FAIL" -Detail $_
    exit 1
  }
}

# Decide whether API tests can run
if (-not $SkipApiTests) {
  if (-not $TenantId -or -not $AppId) {
    Write-TestResult -TestName "API test prerequisites" -Result "WARN" `
    -Detail "TenantId or AppId missing — API tests will be skipped."
    $SkipApiTests = $true
  }
  elseif (-not $ClientSecret -and -not $CertThumbprint) {
    Write-TestResult -TestName "API test prerequisites" -Result "WARN" `
    -Detail "No credential supplied — API tests will be skipped."
    $SkipApiTests = $true
  }
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
        -Detail "Connected successfully."
    }
    catch {
        Write-TestResult -TestName "Exchange Online session" -Result "FAIL" -Detail $_
        exit 1
    }
}

#endregion

#region ── Validate Out-of-Scope Mailbox ───────────────────────────────────────

if ($OutOfScopeMailbox) {
  Write-Section "Validating out-of-scope mailbox"

  # ── Check 1: Mailbox exists in EXO ─────────────────────────────────────
  $outOfScopeMailboxObj = $null
  try {
    $outOfScopeMailboxObj = Get-Mailbox -Identity $OutOfScopeMailbox -ErrorAction Stop
    Write-TestResult `
    -TestName "Out-of-scope mailbox exists: $OutOfScopeMailbox" `
    -Result   "PASS" `
    -Detail   "Mailbox found (Type: $($outOfScopeMailboxObj.RecipientTypeDetails))."
  }
  catch {
    Write-TestResult `
    -TestName "Out-of-scope mailbox exists: $OutOfScopeMailbox" `
    -Result   "WARN" `
    -Detail   ("No EXO mailbox found for '$OutOfScopeMailbox' — " +
      "all out-of-scope tests will be skipped. " +
      "Provide a UPN that has an EXO mailbox but is not " +
    "stamped with $CustomAttribute='$AttrValue'.")
    $OutOfScopeMailbox = $null
  }

  # ── Check 2: Mailbox must NOT have the scope attribute stamped ──────────
  # This is the ground-truth check. If the attribute is present, the mailbox
  # is actually in scope and the out-of-scope tests would produce false results.
  if ($OutOfScopeMailbox -and $outOfScopeMailboxObj) {
    $currentAttrValue = $outOfScopeMailboxObj.$CustomAttribute

    if ($currentAttrValue -eq $AttrValue) {
      Write-TestResult `
      -TestName "Out-of-scope attribute absent: $OutOfScopeMailbox" `
      -Result   "FAIL" `
      -Detail   ("Mailbox has $CustomAttribute='$currentAttrValue' — " +
        "it IS in scope for this app. " +
        "Choose a mailbox that has not been stamped with " +
        "$CustomAttribute='$AttrValue', or remove the stamp " +
      "with: Set-Mailbox '$OutOfScopeMailbox' -$CustomAttribute `$null")
      # Suppress all out-of-scope tests — they would give false results
      $OutOfScopeMailbox = $null
    }
    elseif (-not [string]::IsNullOrWhiteSpace($currentAttrValue)) {
      # Has a different attribute value — still out of scope, but worth noting
      Write-TestResult `
      -TestName "Out-of-scope attribute absent: $OutOfScopeMailbox" `
      -Result   "PASS" `
      -Detail   ("$CustomAttribute='$currentAttrValue' " +
      "(different value — correctly out of scope).")
    }
    else {
      Write-TestResult `
      -TestName "Out-of-scope attribute absent: $OutOfScopeMailbox" `
      -Result   "PASS" `
      -Detail   "$CustomAttribute is empty — correctly out of scope."
    }
  }

  # ── Check 3: Confirm none of the in-scope mailboxes were supplied ───────
  # Catches the mistake of passing an in-scope mailbox as the out-of-scope one.
  if ($OutOfScopeMailbox) {
    $resolvedOos = $outOfScopeMailboxObj.PrimarySmtpAddress.ToLower()
    $overlap     = $Mailboxes | Where-Object { $_.ToLower() -eq $resolvedOos }

    if ($overlap) {
      Write-TestResult `
      -TestName "Out-of-scope mailbox is not in -Mailboxes list" `
      -Result   "FAIL" `
      -Detail   ("'$OutOfScopeMailbox' was also supplied in -Mailboxes. " +
        "The out-of-scope mailbox must be a different mailbox " +
      "that is not configured for this app.")
      $OutOfScopeMailbox = $null
    }
    else {
      Write-TestResult `
      -TestName "Out-of-scope mailbox is not in -Mailboxes list" `
      -Result   "PASS" `
      -Detail   "No overlap with in-scope mailbox list."
    }
  }

  # ── Summary ─────────────────────────────────────────────────────────────
  if ($OutOfScopeMailbox) {
    Write-TestResult `
    -TestName "Out-of-scope mailbox ready" `
    -Result   "INFO" `
    -Detail   ("'$OutOfScopeMailbox' will be used for denial tests. " +
      "Note: Test-ServicePrincipalAuthorization results are " +
    "informational only — Graph API responses are authoritative.")
  }
  else {
    Write-TestResult `
    -TestName "Out-of-scope mailbox ready" `
    -Result   "WARN" `
    -Detail   "Out-of-scope mailbox was disqualified — denial tests will be skipped."
  }
}

#endregion


#region ── RBAC: Management Scope ─────────────────────────────────────────────

Write-Section "RBAC: Management Scope"

Invoke-TestBlock -Name "Scope exists: $ScopeName" -Test {
  $scope = Get-ManagementScope -Identity $ScopeName -ErrorAction Stop
  "Filter: $($scope.RecipientFilter)"
}

Invoke-TestBlock -Name "Scope filter is correct" -Test {
  $scope = Get-ManagementScope -Identity $ScopeName -ErrorAction Stop

  if ($scope.RecipientFilter -notmatch [regex]::Escape($CustomAttribute)) {
    throw ("Filter does not reference '$CustomAttribute'. " +
    "Got: $($scope.RecipientFilter)")
  }
  if ($scope.RecipientFilter -notmatch [regex]::Escape($AttrValue)) {
    throw ("Filter does not reference '$AttrValue'. " +
    "Got: $($scope.RecipientFilter)")
  }
  "Filter OK: $($scope.RecipientFilter)"
}

#endregion

#region ── RBAC: Role Assignments ─────────────────────────────────────────────

Write-Section "RBAC: Role Assignments"

foreach ($roleShort in @("MailRead", "MailSend")) {
  $raName   = "$AppName-$roleShort"
  $roleName = if ($roleShort -eq "MailRead") {
    "Application Mail.Read"
  } else {
    "Application Mail.Send"
  }

  Invoke-TestBlock -Name "Role assignment exists: $raName" -Test {
    $ra = Get-ManagementRoleAssignment -Identity $raName -ErrorAction Stop
    # CustomResourceScope is the correct property name on this object type
    "Role=$($ra.Role)  Scope=$($ra.CustomResourceScope)"
  }

  Invoke-TestBlock -Name "Role assignment scope: $raName → $ScopeName" -Test {
    $ra = Get-ManagementRoleAssignment -Identity $raName -ErrorAction Stop
    if ($ra.CustomResourceScope -ne $ScopeName) {
      throw ("Expected scope '$ScopeName', " +
      "got '$($ra.CustomResourceScope)'.")
    }
    "Scope matches."
  }
}

#endregion

#region ── RBAC: EXO Service Principal ────────────────────────────────────────

Write-Section "RBAC: EXO Service Principal"

Invoke-TestBlock -Name "EXO service principal exists: $AppName" -Test {
  $exoSp = Get-ServicePrincipal -ErrorAction SilentlyContinue |
  Where-Object { $_.DisplayName -eq $AppName } |
  Select-Object -First 1

  if (-not $exoSp) {
    throw "EXO service principal '$AppName' not found."
  }
  "AppId: $($exoSp.AppId)"
}

#endregion

#region ── RBAC: Mail-Enabled Security Group ──────────────────────────────────

Write-Section "RBAC: Mail-Enabled Security Group"

Invoke-TestBlock -Name "MESG exists: $MESGName" -Test {
  $mesg    = Get-DistributionGroup -Identity $MESGName -ErrorAction Stop
  $members = Get-DistributionGroupMember -Identity $MESGName
  "Members: $($members.Count)"
}

foreach ($addr in $Mailboxes) {
  Invoke-TestBlock -Name "MESG membership: $addr" -Test {
    $members = Get-DistributionGroupMember -Identity $MESGName -ErrorAction Stop
    $match   = $members | Where-Object { $_.PrimarySmtpAddress -eq $addr }
    if (-not $match) {
      throw "'$addr' is not a member of '$MESGName'."
    }
    "Member confirmed."
  }
}

#endregion

#region ── RBAC: Mailbox Custom Attributes ────────────────────────────────────

Write-Section "RBAC: Mailbox Custom Attributes"

foreach ($addr in $Mailboxes) {
  Invoke-TestBlock -Name "Attribute stamp: $addr → $CustomAttribute = '$AttrValue'" -Test {
    $mbx   = Get-Mailbox -Identity $addr -ErrorAction Stop
    $value = $mbx.$CustomAttribute
    if ($value -ne $AttrValue) {
      throw "Expected '$AttrValue', got '$value'."
    }
    "Attribute verified."
  }
}

#endregion

#region ── RBAC: Scope Authorization Check ────────────────────────────────────

Write-Section "RBAC: Scope Authorization Check"

# In-scope mailboxes — these must return InScope=True for all roles
foreach ($addr in $Mailboxes) {
  Invoke-TestBlock -Name "In-scope: $addr" -Test {
    $results = Test-ServicePrincipalAuthorization `
    -Identity $AppName `
    -Resource $addr `
    -ErrorAction Stop

    $outOfScope = $results | Where-Object { -not $_.InScope }
    if ($outOfScope) {
      $outRoles = ($outOfScope | Select-Object -ExpandProperty RoleName) -join ", "
      throw "Not in-scope for roles: $outRoles"
    }
    "All roles in-scope."
  }
}

# Out-of-scope mailbox — Test-ServicePrincipalAuthorization is NOT a reliable
# deny check for non-exclusive management scopes. The cmdlet reports whether
# the role assignment's scope filter matches the mailbox, but due to how EXO
# evaluates non-exclusive scopes internally it can return InScope=True for
# mailboxes that the app cannot actually access via the Graph API.
#
# The authoritative deny check is the Graph API response (403/404), which is
# tested in the Graph API section below.
#
# Here we only verify the attribute is absent (ground truth for scope
# exclusion) and emit the Test-ServicePrincipalAuthorization result as
# informational so the operator can see what EXO reports without it
# counting as a test failure.
if ($OutOfScopeMailbox) {

  Invoke-TestBlock -Name "Out-of-scope attribute absent: $OutOfScopeMailbox" -Test {
    $mbx   = Get-Mailbox -Identity $OutOfScopeMailbox -ErrorAction Stop
    $value = $mbx.$CustomAttribute
    if ($value -eq $AttrValue) {
      throw ("Mailbox has $CustomAttribute='$value' — it IS in scope. " +
        "Choose a mailbox not stamped with $CustomAttribute='$AttrValue', " +
        "or remove the stamp with: " +
      "Set-Mailbox '$OutOfScopeMailbox' -$CustomAttribute `$null")
    }
    "Confirmed: $CustomAttribute='$value' (not '$AttrValue') — not in scope."
  }

  # Informational only — result is not recorded as PASS or FAIL
  Write-Host ""
  Write-Host "  ℹ  [INFO] Test-ServicePrincipalAuthorization (informational only): $OutOfScopeMailbox" `
  -ForegroundColor DarkCyan

  try {
    $authResults = Test-ServicePrincipalAuthorization `
    -Identity    $AppName `
    -Resource    $OutOfScopeMailbox `
    -ErrorAction Stop

    foreach ($r in $authResults) {
      $detail = "  ℹ    Role: $($r.RoleName)  InScope: $($r.InScope)  " +
      "Scope: $($r.AllowedResourceScope)"
      Write-Host $detail -ForegroundColor DarkCyan
    }

    Write-Host ("  ℹ  Note: InScope=True here does not mean access is granted. " +
      "Non-exclusive management scopes are not evaluated as deny rules " +
    "by this cmdlet. Graph API responses are the authoritative test.") `
    -ForegroundColor DarkCyan
  }
  catch {
    Write-Host "  ℹ  Test-ServicePrincipalAuthorization could not be run: $_" `
    -ForegroundColor DarkCyan
  }
}

#endregion

#region ── Graph API: Token Acquisition ───────────────────────────────────────

if (-not $SkipApiTests) {
  Write-Section "Graph API: Token Acquisition"

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

      $toSign    = "$header.$payload"
      $rsa       = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
      $sigBytes  = $rsa.SignData(
        [Text.Encoding]::UTF8.GetBytes($toSign),
        [Security.Cryptography.HashAlgorithmName]::SHA256,
        [Security.Cryptography.RSASignaturePadding]::Pkcs1
      )
      $sig       = [Convert]::ToBase64String($sigBytes).TrimEnd('=').Replace('+','-').Replace('/','_')
      $assertion = "$toSign.$sig"

      $tokenBody = @{
        client_id             = $AppId
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        client_assertion      = $assertion
        scope                 = "https://graph.microsoft.com/.default"
        grant_type            = "client_credentials"
      }
    }
    else {
      $tokenBody = @{
        client_id     = $AppId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
      }
    }

    $tokenResponse = Invoke-RestMethod `
    -Method      Post `
    -Uri         "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -Body        $tokenBody `
    -ErrorAction Stop

    Assert-NotNull -Value $tokenResponse.access_token -Label "Access token"
    $script:accessToken = $tokenResponse.access_token

    $expiry = (Get-Date).AddSeconds($tokenResponse.expires_in)
    "Token acquired. Expires: $($expiry.ToString('HH:mm:ss'))"
  }
}

#endregion

#region ── Graph API: Mail Read Tests ─────────────────────────────────────────

if (-not $SkipApiTests -and $script:accessToken) {

  $headers = @{
    Authorization  = "Bearer $script:accessToken"
    "Content-Type" = "application/json"
  }

  Write-Section "Graph API: Mail Read (in-scope mailboxes — expect 200)"

  foreach ($addr in $Mailboxes) {
    Invoke-TestBlock -Name "Read mail: $addr" -Test {
      $uri      = "https://graph.microsoft.com/v1.0/users/$addr/messages" +
      "?`$top=1&`$select=id,subject"
      $response = Invoke-RestMethod `
      -Headers     $headers `
      -Uri         $uri `
      -Method      Get `
      -ErrorAction Stop
      $count = ($response.value | Measure-Object).Count
      "Retrieved $count message(s)."
    }
  }

  if ($OutOfScopeMailbox) {
    Write-Section "Graph API: Mail Read (out-of-scope — expect 403/404)"

    Invoke-TestBlock `
    -Name            "Read denied: $OutOfScopeMailbox" `
    -ExpectedFailure "403|404|Forbidden|Not Found|ErrorAccessDenied" `
    -Test {
      $uri = "https://graph.microsoft.com/v1.0/users/$OutOfScopeMailbox/messages" +
      "?`$top=1"
      Invoke-RestMethod `
      -Headers     $headers `
      -Uri         $uri `
      -Method      Get `
      -ErrorAction Stop
    }
  }

  #endregion

  #region ── Graph API: Mail Send Tests ─────────────────────────────────────────

  Write-Section "Graph API: Mail Send (in-scope mailboxes — expect 202)"

  foreach ($addr in $Mailboxes) {
    $recipient = if ($TestRecipient) { $TestRecipient } else { $addr }

    Invoke-TestBlock -Name "Send mail: $addr → $recipient" -Test {
      $uri  = "https://graph.microsoft.com/v1.0/users/$addr/sendMail"
      $body = @{
        message = @{
          subject      = "[Test-ScopedMailboxApp] Scoped send test — $(Get-Date -Format 'u')"
          body         = @{
            contentType = "Text"
            content     = "Automated test from Test-ScopedMailboxApp.ps1."
          }
          toRecipients = @(
            @{ emailAddress = @{ address = $recipient } }
          )
        }
        saveToSentItems = $false
      } | ConvertTo-Json -Depth 10

      Invoke-RestMethod `
      -Headers     $headers `
      -Uri         $uri `
      -Method      Post `
      -Body        $body `
      -ErrorAction Stop
      "Message sent to $recipient."
    }
  }

  if ($OutOfScopeMailbox) {
    Write-Section "Graph API: Mail Send (out-of-scope — expect 403/404)"

    $recipient = if ($TestRecipient) { $TestRecipient } else { $Mailboxes[0] }

    Invoke-TestBlock `
    -Name            "Send denied: $OutOfScopeMailbox" `
    -ExpectedFailure "403|404|Forbidden|Not Found|ErrorAccessDenied" `
    -Test {
      $uri  = "https://graph.microsoft.com/v1.0/users/$OutOfScopeMailbox/sendMail"
      $body = @{
        message = @{
          subject      = "[Test] Should be denied"
          body         = @{ contentType = "Text"; content = "Should not arrive." }
          toRecipients = @(@{ emailAddress = @{ address = $recipient } })
        }
      } | ConvertTo-Json -Depth 10

      Invoke-RestMethod `
      -Headers     $headers `
      -Uri         $uri `
      -Method      Post `
      -Body        $body `
      -ErrorAction Stop
    }
  }
}
elseif ($SkipApiTests) {
  Write-Section "Graph API Tests"
  Write-TestResult -TestName "Graph API tests" -Result "SKIP" `
  -Detail "Skipped via -SkipApiTests or missing credentials."
}

#endregion

#region ── Final Results ───────────────────────────────────────────────────────

Write-Host ""
Write-Host ("═" * 70) -ForegroundColor DarkGray
Write-Host "  TEST RESULTS" -ForegroundColor White
Write-Host ("═" * 70) -ForegroundColor DarkGray

# ── Counts ─────────────────────────────────────────────────────────────────
Write-Host ("  ✔  Passed  : {0,3}" -f $script:TestsPassed.Count) -ForegroundColor Green
Write-Host ("  ✘  Failed  : {0,3}" -f $script:TestsFailed.Count) -ForegroundColor Red
Write-Host ("  ⚠  Warned  : {0,3}" -f $script:TestsWarned.Count) -ForegroundColor Yellow

# ── Per-category breakdown ──────────────────────────────────────────────────
$categories = [ordered]@{
    "RBAC"       = @{
        Passed = $script:TestsPassed | Where-Object { $_ -match "^(Scope|Role|EXO|MESG|Attribute|In-scope|Out-of-scope attr)" }
        Failed = $script:TestsFailed | Where-Object { $_ -match "^(Scope|Role|EXO|MESG|Attribute|In-scope|Out-of-scope attr)" }
        Warned = $script:TestsWarned | Where-Object { $_ -match "^(Scope|Role|EXO|MESG|Attribute|In-scope|Out-of-scope attr)" }
    }
    "Graph API"  = @{
        Passed = $script:TestsPassed | Where-Object { $_ -match "^(Read mail|Send mail|Read denied|Send denied|Acquire access token)" }
        Failed = $script:TestsFailed | Where-Object { $_ -match "^(Read mail|Send mail|Read denied|Send denied|Acquire access token)" }
        Warned = $script:TestsWarned | Where-Object { $_ -match "^(Read mail|Send mail|Read denied|Send denied|Acquire access token)" }
    }
}

Write-Host ""
Write-Host "  Breakdown by category:" -ForegroundColor White
Write-Host ("  {0,-12} {1,8} {2,8} {3,8}" -f "Category", "Passed", "Failed", "Warned") `
    -ForegroundColor DarkGray
Write-Host ("  {0,-12} {1,8} {2,8} {3,8}" -f "────────────", "──────", "──────", "──────") `
    -ForegroundColor DarkGray

foreach ($cat in $categories.Keys) {
    $p = @($categories[$cat].Passed).Count
    $f = @($categories[$cat].Failed).Count
    $w = @($categories[$cat].Warned).Count

    $rowColor = if     ($f -gt 0) { "Red"     }
                elseif ($w -gt 0) { "Yellow"  }
                elseif ($p -gt 0) { "Green"   }
                else              { "DarkGray" }

    $skipped  = if ($p -eq 0 -and $f -eq 0 -and $w -eq 0) { " (skipped)" } else { "" }

    Write-Host ("  {0,-12} {1,8} {2,8} {3,8}{4}" -f $cat, $p, $f, $w, $skipped) `
        -ForegroundColor $rowColor
}

# ── Failed test detail ──────────────────────────────────────────────────────
if ($script:TestsFailed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Failed tests:" -ForegroundColor Red
    $script:TestsFailed | ForEach-Object {
        Write-Host "    • $_" -ForegroundColor Red
    }
}

# ── Warning detail ──────────────────────────────────────────────────────────
if ($script:TestsWarned.Count -gt 0) {
    Write-Host ""
    Write-Host "  Warnings:" -ForegroundColor Yellow
    $script:TestsWarned | ForEach-Object {
        Write-Host "    • $_" -ForegroundColor Yellow
    }
}

# ── Overall verdict ─────────────────────────────────────────────────────────
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

