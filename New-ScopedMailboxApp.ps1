#Requires -Version 5.1
<#
    .SYNOPSIS
    Creates an Entra ID app registration with scoped Exchange Online mailbox
    read and send access, restricted to a specified list of mailboxes.

    .DESCRIPTION
    This script:
    - Creates an Entra ID app registration and service principal
    - Creates a client secret (or optionally a certificate)
    - Creates a mail-enabled security group and adds specified mailboxes
    - Stamps a custom attribute on each mailbox for scope filtering
    - Creates an Exchange Online management scope
    - Registers the EXO service principal
    - Assigns scoped Application Mail.Read and Mail.Send roles
    - Verifies each step before proceeding

    .PARAMETER AppName
    Display name for the app registration and associated resources.
    Example: "MailApp-HR"

    .PARAMETER Mailboxes
    Array of mailbox SMTP addresses to scope access to.
    Example: @("hr@contoso.com", "payroll@contoso.com")

    .PARAMETER UseCertificate
    If specified, creates a self-signed certificate credential instead of
    a client secret.

    .PARAMETER CustomAttribute
    The mailbox CustomAttribute (1-15) to use for scope filtering.
    Defaults to CustomAttribute1.

    .PARAMETER SecretExpiryMonths
    Number of months until the client secret expires. Defaults to 6.

    .PARAMETER OutputPath
    Path to write the configuration summary JSON file.
    Defaults to the current directory.

    .EXAMPLE
    .\New-ScopedMailboxApp.ps1 `
    -AppName "MailApp-HR" `
    -Mailboxes @("hr@contoso.com", "payroll@contoso.com")

    .EXAMPLE
    .\New-ScopedMailboxApp.ps1 `
    -AppName "MailApp-Finance" `
    -Mailboxes @("finance@contoso.com") `
    -UseCertificate `
    -CustomAttribute "CustomAttribute2"

    .NOTES
    Requires:
    - Microsoft.Graph PowerShell module
    - ExchangeOnlineManagement module
    - Exchange Online global admin or equivalent
    - Entra ID Application Administrator or higher
#>

[CmdletBinding(SupportsShouldProcess)]
param (
  [Parameter(Mandatory)]
  [ValidateNotNullOrEmpty()]
  [string]$AppName,

  [Parameter(Mandatory)]
  [ValidateNotNullOrEmpty()]
  [string[]]$Mailboxes,

  [Parameter()]
  [switch]$UseCertificate,

  [Parameter()]
  [ValidateSet(
      "CustomAttribute1",  "CustomAttribute2",  "CustomAttribute3",
      "CustomAttribute4",  "CustomAttribute5",  "CustomAttribute6",
      "CustomAttribute7",  "CustomAttribute8",  "CustomAttribute9",
      "CustomAttribute10", "CustomAttribute11", "CustomAttribute12",
      "CustomAttribute13", "CustomAttribute14", "CustomAttribute15"
  )]
  [string]$CustomAttribute = "CustomAttribute15",

  [Parameter()]
  [ValidateRange(1, 24)]
  [int]$SecretExpiryMonths = 6,

  [Parameter()]
  [ValidateNotNullOrEmpty()]
  [string]$OutputPath = (Get-Location).Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region ── Helpers ─────────────────────────────────────────────────────────────

$script:StepNumber  = 0
$script:StepsFailed = @()

# Cross-step state — all declared at script scope so Invoke-Step blocks can write them
$script:app               = $null
$script:sp                = $null
$script:exoSp             = $null
$script:validatedMailboxes = [System.Collections.Generic.List[object]]::new()
$script:config            = $null

# ── Compatible module versions ───────────────────────────────────────────────
# ExchangeOnlineManagement 3.8+ and Microsoft.Graph 2.25+ cannot coexist in
# the same PowerShell session due to a WAM/MSAL broker DLL conflict.
# These versions are the last known-good combination where both load cleanly.
$script:RequiredExoVersion   = "3.7.0"
$script:RequiredGraphVersion = "2.24.0"

function Write-Step {
  param([string]$Message)
  $script:StepNumber++
  Write-Host ""
  Write-Host ("─" * 70) -ForegroundColor DarkGray
  Write-Host "  STEP $($script:StepNumber) │ $Message" -ForegroundColor Cyan
  Write-Host ("─" * 70) -ForegroundColor DarkGray
}

function Write-Success {
  param([string]$Message)
  Write-Host "  ✔  $Message" -ForegroundColor Green
}

function Write-Info {
  param([string]$Message)
  Write-Host "  ℹ  $Message" -ForegroundColor DarkCyan
}

function Write-Warn {
  param([string]$Message)
  Write-Host "  ⚠  $Message" -ForegroundColor Yellow
}

function Assert-NotNull {
  param(
    [object]$Value,
    [string]$Label
  )
  if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) {
    throw "$Label was null or empty — cannot continue."
  }
}

function Invoke-Step {
  param(
    [string]$Name,
    [scriptblock]$Action,
    [switch]$ContinueOnError
  )

  try {
    & $Action
    Write-Success "$Name completed."
  }
  catch {
    $script:StepsFailed += $Name
    Write-Host "  ✘  $Name FAILED: $_" -ForegroundColor Red

    if (-not $ContinueOnError) {
      Write-Host ""
      Write-Host "  ► Aborting. Run the script again after resolving the issue." `
      -ForegroundColor Red
      Write-SummaryAndExit -ExitCode 1
    }
    else {
      Write-Warn "Continuing despite error (ContinueOnError set)."
    }
  }
}

function Write-SummaryAndExit {
  param([int]$ExitCode = 0)

  Write-Host ""
  Write-Host ("═" * 70) -ForegroundColor DarkGray
  Write-Host "  SUMMARY" -ForegroundColor White
  Write-Host ("═" * 70) -ForegroundColor DarkGray

  if ($script:StepsFailed.Count -eq 0) {
    Write-Host "  All steps completed successfully." -ForegroundColor Green
  }
  else {
    Write-Host "  Failed steps:" -ForegroundColor Red
    $script:StepsFailed | ForEach-Object {
      Write-Host "    • $_" -ForegroundColor Red
    }
  }

  Write-Host ("═" * 70) -ForegroundColor DarkGray
  exit $ExitCode
}

#endregion

#region ── Banner ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("═" * 70) -ForegroundColor DarkCyan
Write-Host "  New-ScopedMailboxApp  │  Entra + Exchange Online Configuration" `
-ForegroundColor Cyan
Write-Host ("═" * 70) -ForegroundColor DarkCyan
Write-Host "  App Name       : $AppName"
Write-Host "  Mailboxes      : $($Mailboxes -join ', ')"
Write-Host "  Credential     : $(if ($UseCertificate) { 'Certificate' } else { 'Client Secret' })"
Write-Host "  Scope Attr     : $CustomAttribute"
Write-Host "  Output Path    : $OutputPath"
Write-Host ("═" * 70) -ForegroundColor DarkCyan

#endregion

#region ── Derived Names ───────────────────────────────────────────────────────

$MESGName    = "$AppName-MESG"
$MESGAlias   = ($AppName -replace '[^a-zA-Z0-9]', '') + "MESG"
$ScopeName   = "$AppName-Scope"
$AttrValue   = $AppName
$ScopeFilter = "$CustomAttribute -eq '$AttrValue'"

$script:config = [ordered]@{
  AppName         = $AppName
  Mailboxes       = $Mailboxes
  CustomAttribute = $CustomAttribute
  AttributeValue  = $AttrValue
  ScopeFilter     = $ScopeFilter
  MESGName        = $MESGName
  ScopeName       = $ScopeName
  TenantId        = $null
  AppId           = $null
  AppObjectId     = $null
  SpObjectId      = $null
  CredentialType  = if ($UseCertificate) { "Certificate" } else { "ClientSecret" }
  ClientSecret    = $null
  CertThumbprint  = $null
  CreatedAt       = (Get-Date -Format "o")
}

#endregion

#region ── Step 0: Prerequisites ───────────────────────────────────────────────

Write-Step "Checking prerequisites"

Invoke-Step -Name "Module check" -Action {
  $required = @(
    @{ Name = "ExchangeOnlineManagement";    Version = $script:RequiredExoVersion   }
    @{ Name = "Microsoft.Graph.Authentication"; Version = $script:RequiredGraphVersion }
    @{ Name = "Microsoft.Graph.Applications";  Version = $script:RequiredGraphVersion }
  )

  foreach ($mod in $required) {
    $installed = Get-InstalledModule `
    -Name            $mod.Name `
    -RequiredVersion $mod.Version `
    -ErrorAction     SilentlyContinue

    if (-not $installed) {
      throw (
        "Module '$($mod.Name)' version $($mod.Version) is not installed.`n" +
        "Run: Install-Module $($mod.Name) -RequiredVersion $($mod.Version) " +
        "-Scope CurrentUser -Force -AllowClobber"
      )
    }
    Write-Info "Module OK: $($mod.Name) $($mod.Version)"
  }
}

Invoke-Step -Name "Output path check" -Action {
  if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    Write-Info "Created output directory: $OutputPath"
  }
  else {
    Write-Info "Output directory exists: $OutputPath"
  }
}

#endregion

#region ── Step 1: Connect to Microsoft Graph ──────────────────────────────────

Write-Step "Connecting to Microsoft Graph"

Invoke-Step -Name "Graph connection" -Action {
  Import-Module Microsoft.Graph.Authentication `
  -RequiredVersion $script:RequiredGraphVersion `
  -ErrorAction Stop

  Import-Module Microsoft.Graph.Applications `
  -RequiredVersion $script:RequiredGraphVersion `
  -ErrorAction Stop

  Connect-MgGraph -Scopes @(
    "Application.ReadWrite.All",
    "Directory.ReadWrite.All"
  ) -NoWelcome

  $ctx = Get-MgContext
  Assert-NotNull -Value $ctx          -Label "Graph context"
  Assert-NotNull -Value $ctx.TenantId -Label "Tenant ID"

  $script:config.TenantId = $ctx.TenantId
  Write-Info "Connected to tenant: $($ctx.TenantId)"
  Write-Info "Signed in as: $($ctx.Account)"
}

#endregion

#region ── Step 2: Connect to Exchange Online ──────────────────────────────────

Write-Step "Connecting to Exchange Online"

Invoke-Step -Name "Exchange Online connection" -Action {
  Import-Module ExchangeOnlineManagement `
  -RequiredVersion $script:RequiredExoVersion `
  -ErrorAction Stop

  try {
    $null = Get-OrganizationConfig -ErrorAction Stop
    Write-Info "Already connected to Exchange Online."
  }
  catch {
    Connect-ExchangeOnline `
    -ShowBanner:$false `
    -ErrorAction Stop

    $null = Get-OrganizationConfig -ErrorAction Stop
    Write-Info "Exchange Online connection verified."
  }
}

#endregion

#region ── Step 2b: Check AllowServicePrincipalSmtpAuth ───────────────────────

Write-Step "Checking org-level SMTP AUTH settings"

Invoke-Step -Name "Org SMTP AUTH check" -ContinueOnError -Action {
  $transport = Get-TransportConfig -ErrorAction Stop

  # SmtpClientAuthenticationDisabled — must be False for any SMTP AUTH
  $orgDisabled = $transport.SmtpClientAuthenticationDisabled
  if ($orgDisabled -eq $true) {
    Write-Warn ("SmtpClientAuthenticationDisabled = True at org level. " +
      "SMTP AUTH is globally disabled. " +
    "Fix: Set-TransportConfig -SmtpClientAuthenticationDisabled `$false")
  }
  else {
    Write-Info "SmtpClientAuthenticationDisabled = $orgDisabled (SMTP AUTH not globally off)."
  }

  # AllowServicePrincipalSmtpAuth — required for SMTP.SendAsApp client credentials flow.
  # This property may not exist on all tenants/module versions — handle gracefully.
  $spSmtp = $transport.PSObject.Properties["AllowServicePrincipalSmtpAuth"]
  if ($null -eq $spSmtp) {
    Write-Warn ("AllowServicePrincipalSmtpAuth property not present on this tenant. " +
    "If SMTP sending fails with 535, contact Microsoft Support.")
  }
  elseif ($spSmtp.Value -ne $true) {
    Write-Warn ("AllowServicePrincipalSmtpAuth = $($spSmtp.Value). " +
      "This must be True for the SMTP.SendAsApp client credentials flow. " +
    "Fix: Set-TransportConfig -AllowServicePrincipalSmtpAuth `$true")
  }
  else {
    Write-Info "AllowServicePrincipalSmtpAuth = $($spSmtp.Value)."
  }
}

#endregion

#region ── Step 3: Validate Mailboxes ─────────────────────────────────────────

Write-Step "Validating mailbox addresses"

Invoke-Step -Name "Mailbox validation" -Action {
  foreach ($addr in $Mailboxes) {
    try {
      $mbx = Get-Mailbox -Identity $addr -ErrorAction Stop
      $script:validatedMailboxes.Add($mbx)
      Write-Info "Validated: $addr  (Type: $($mbx.RecipientTypeDetails))"
    }
    catch {
      throw "Mailbox '$addr' not found or inaccessible: $_"
    }
  }
  Write-Info "$($script:validatedMailboxes.Count) of $($Mailboxes.Count) mailboxes validated."
}

#endregion

#region ── Step 4: Create App Registration ─────────────────────────────────────

Write-Step "Creating Entra ID app registration"

Invoke-Step -Name "App registration" -Action {
  $existing = Get-MgApplication `
  -Filter "displayName eq '$AppName'" `
  -ErrorAction SilentlyContinue |
  Select-Object -First 1

  if ($existing) {
    Write-Warn "App '$AppName' already exists (ID: $($existing.AppId)). Using existing."
    $script:app = $existing
  }
  else {
    # Create and then re-fetch to guarantee a fully hydrated object
    New-MgApplication `
    -DisplayName    $AppName `
    -SignInAudience "AzureADMyOrg" | Out-Null

    Write-Info "App created — waiting for replication..."
    Start-Sleep -Seconds 5

    $script:app = Get-MgApplication `
    -Filter "displayName eq '$AppName'" `
    -ErrorAction Stop |
    Select-Object -First 1
  }

  Assert-NotNull -Value $script:app        -Label "App registration object"
  Assert-NotNull -Value $script:app.AppId  -Label "App AppId"
  Assert-NotNull -Value $script:app.Id     -Label "App Object ID"

  $script:config.AppId       = $script:app.AppId
  $script:config.AppObjectId = $script:app.Id

  Write-Info "App ID (Client ID) : $($script:app.AppId)"
  Write-Info "App Object ID      : $($script:app.Id)"
}

#endregion

#region ── Step 5: Create Service Principal ────────────────────────────────────

Write-Step "Creating Entra ID service principal"

Invoke-Step -Name "Service principal" -Action {
  Assert-NotNull -Value $script:app.AppId -Label "App AppId (pre-SP creation)"

  $existing = Get-MgServicePrincipal `
  -Filter      "appId eq '$($script:app.AppId)'" `
  -ErrorAction SilentlyContinue |
  Select-Object -First 1

  if ($existing) {
    Write-Warn "Service principal already exists. Using existing."
    $script:sp = $existing
  }
  else {
    # Create and immediately re-fetch to guarantee a fully hydrated object.
    # Do NOT use the object returned by New-MgServicePrincipal directly —
    # it may be partially populated depending on module version.
    New-MgServicePrincipal -AppId $script:app.AppId -ErrorAction Stop | Out-Null

    Start-Sleep -Seconds 5

    $script:sp = Get-MgServicePrincipal `
    -Filter      "appId eq '$($script:app.AppId)'" `
    -ErrorAction Stop |
    Select-Object -First 1
  }

  Assert-NotNull -Value $script:sp    -Label "Service principal object"
  Assert-NotNull -Value $script:sp.Id -Label "Service principal Object ID"

  $script:config.SpObjectId = $script:sp.Id

  # Log both IDs explicitly so any mismatch is immediately visible
  Write-Info "App Registration Object ID : $($script:app.Id)   ← do NOT use for EXO"
  Write-Info "Service Principal Object ID: $($script:sp.Id)    ← use this for EXO"
  Write-Info "AppId (Client ID)          : $($script:app.AppId)"
}

#endregion

#region ── Step 6: Create Credential ──────────────────────────────────────────

Write-Step "Creating app credential ($($script:config.CredentialType))"

Invoke-Step -Name "Credential creation" -Action {
  Assert-NotNull -Value $script:app.Id -Label "App Object ID (pre-credential)"

  if ($UseCertificate) {
    $cert = New-SelfSignedCertificate `
    -Subject           "CN=$AppName" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy   Exportable `
    -KeySpec           Signature `
    -KeyLength         2048 `
    -HashAlgorithm     SHA256 `
    -NotAfter          (Get-Date).AddYears(2)

    $certPath = Join-Path $OutputPath "$AppName.cer"
    Export-Certificate -Cert $cert -FilePath $certPath | Out-Null

    $certBytes = [IO.File]::ReadAllBytes($certPath)

    Update-MgApplication -ApplicationId $script:app.Id -KeyCredentials @(@{
        DisplayName = "$AppName-Cert"
        Type        = "AsymmetricX509Cert"
        Usage       = "Verify"
        Key         = $certBytes
    })

    $script:config.CertThumbprint = $cert.Thumbprint
    Write-Info "Certificate thumbprint : $($cert.Thumbprint)"
    Write-Info "Certificate exported to: $certPath"
  }
  else {
    $secretResult = Add-MgApplicationPassword `
    -ApplicationId $script:app.Id `
    -BodyParameter @{
      PasswordCredential = @{
        DisplayName = "$AppName-Secret"
        EndDateTime = (Get-Date).AddMonths($SecretExpiryMonths)
      }
    }

    Assert-NotNull -Value $secretResult.SecretText -Label "Client secret text"
    $script:config.ClientSecret = $secretResult.SecretText

    Write-Info "Client secret created (expires: $((Get-Date).AddMonths($SecretExpiryMonths).ToString('yyyy-MM-dd')))"
    Write-Warn "Client secret will only be shown in the output file — store it securely."
  }
}

#endregion

#region ── Step 7: Create Mail-Enabled Security Group ─────────────────────────

Write-Step "Creating mail-enabled security group: $MESGName"

Invoke-Step -Name "MESG creation" -Action {
  $existingMESG = Get-DistributionGroup -Identity $MESGName -ErrorAction SilentlyContinue
  if ($existingMESG) {
    Write-Warn "MESG '$MESGName' already exists. Using existing."
  }
  else {
    New-DistributionGroup `
    -Name        $MESGName `
    -DisplayName $MESGName `
    -Alias       $MESGAlias `
    -Type        Security `
    -ErrorAction Stop | Out-Null

    Write-Info "MESG created: $MESGName"
    Start-Sleep -Seconds 5
  }
}

Invoke-Step -Name "MESG membership" -Action {
  foreach ($mbx in $script:validatedMailboxes) {
    try {
      Add-DistributionGroupMember `
      -Identity    $MESGName `
      -Member      $mbx.PrimarySmtpAddress `
      -ErrorAction Stop
      Write-Info "Added to MESG: $($mbx.PrimarySmtpAddress)"
    }
    catch {
      if ($_ -match "already a member") {
        Write-Warn "Already a member: $($mbx.PrimarySmtpAddress)"
      }
      else {
        throw $_
      }
    }
  }
}

#endregion

#region ── Step 8: Stamp Custom Attribute on Mailboxes ────────────────────────

Write-Step "Stamping $CustomAttribute = '$AttrValue' on mailboxes"

Invoke-Step -Name "Mailbox attribute stamping" -Action {
  foreach ($mbx in $script:validatedMailboxes) {
    $params = @{
      Identity         = $mbx.PrimarySmtpAddress
      $CustomAttribute = $AttrValue
      ErrorAction      = "Stop"
    }
    Set-Mailbox @params
    Write-Info "Stamped: $($mbx.PrimarySmtpAddress)"
  }

  Write-Info "Verifying attribute stamps..."
  foreach ($mbx in $script:validatedMailboxes) {
    $check = Get-Mailbox -Identity $mbx.PrimarySmtpAddress |
    Select-Object -ExpandProperty $CustomAttribute
    if ($check -ne $AttrValue) {
      throw "Attribute verification failed for $($mbx.PrimarySmtpAddress). " +
      "Expected '$AttrValue', got '$check'."
    }
    Write-Info "Verified: $($mbx.PrimarySmtpAddress) → $CustomAttribute = '$check'"
  }
}

#endregion

#region ── Step 9: Create Management Scope ────────────────────────────────────

Write-Step "Creating Exchange management scope: $ScopeName"

Invoke-Step -Name "Management scope" -Action {
  $existingScope = Get-ManagementScope -Identity $ScopeName -ErrorAction SilentlyContinue
  if ($existingScope) {
    Write-Warn "Scope '$ScopeName' already exists. Using existing."
  }
  else {
    New-ManagementScope `
    -Name                       $ScopeName `
    -RecipientRestrictionFilter $ScopeFilter `
    -ErrorAction                Stop | Out-Null

    Write-Info "Scope created with filter: $ScopeFilter"
  }

  $verifyScope = Get-ManagementScope -Identity $ScopeName -ErrorAction Stop
  Assert-NotNull -Value $verifyScope -Label "Management scope"
  Write-Info "Scope verified : $($verifyScope.Name)"
  Write-Info "Scope filter   : $($verifyScope.RecipientFilter)"
}

#endregion

#region ── Step 10: Register EXO Service Principal ────────────────────────────

Write-Step "Registering Exchange Online service principal"

Invoke-Step -Name "EXO service principal registration" -Action {
  Assert-NotNull -Value $script:app.AppId -Label "App AppId (pre-EXO SP)"
  Assert-NotNull -Value $script:sp.Id     -Label "SP Object ID (pre-EXO SP)"

  # Check if already registered
  $script:exoSp = Get-ServicePrincipal -ErrorAction SilentlyContinue |
  Where-Object { $_.AppId -eq $script:app.AppId } |
  Select-Object -First 1

  if ($script:exoSp) {
    Write-Warn "EXO service principal already exists. Using existing."
  }
  else {
    # Retry loop — guards against residual replication lag beyond the
    # polling window in Step 5.
    $maxAttempts  = 5
    $retryWait    = 15
    $attempt      = 0
    $lastError    = $null

    while ($attempt -lt $maxAttempts -and -not $script:exoSp) {
      $attempt++
      try {
        Write-Info "Registering EXO service principal (attempt $attempt of $maxAttempts)..."

        $script:exoSp = New-ServicePrincipal `
        -AppId       $script:app.AppId `
        -ObjectId    $script:sp.Id `
        -DisplayName $AppName `
        -ErrorAction Stop

        Write-Info "EXO service principal created on attempt $attempt."
      }
      catch {
        $lastError = $_
        if ($attempt -lt $maxAttempts) {
          Write-Warn ("Attempt $attempt failed: $lastError")
          Write-Info "Waiting ${retryWait}s before retry..."
          Start-Sleep -Seconds $retryWait
        }
      }
    }

    if (-not $script:exoSp) {
      throw ("Failed to register EXO service principal after $maxAttempts attempts. " +
      "Last error: $lastError")
    }
  }

  Assert-NotNull -Value $script:exoSp          -Label "EXO service principal"
  Assert-NotNull -Value $script:exoSp.Identity -Label "EXO SP Identity"
  Write-Info "EXO SP Identity: $($script:exoSp.Identity)"
}

#endregion

#region ── Step 11: Assign Scoped Roles ───────────────────────────────────────

Write-Step "Assigning scoped management role assignments"

foreach ($roleSpec in @(
    @{ Short = "MailRead"; Role = "Application Mail.Read" }
    @{ Short = "MailSend"; Role = "Application Mail.Send" }
)) {
  Invoke-Step -Name "$($roleSpec.Role) role assignment" -Action {
    $raName   = "$AppName-$($roleSpec.Short)"
    $existing = Get-ManagementRoleAssignment -Identity $raName -ErrorAction SilentlyContinue

    if ($existing) {
      Write-Warn "Role assignment '$raName' already exists."

      # Verify it has the correct scope even if pre-existing
      if ([string]::IsNullOrWhiteSpace($existing.CustomResourceScope)) {
        throw ("Existing assignment '$raName' has no CustomResourceScope — " +
        "it is tenant-wide. Remove it and re-run the script.")
      }
      if ($existing.CustomResourceScope -ne $ScopeName) {
        throw ("Existing assignment '$raName' has scope " +
          "'$($existing.CustomResourceScope)' instead of '$ScopeName'. " +
        "Remove it and re-run the script.")
      }
      Write-Info "Existing assignment scope verified: $($existing.CustomResourceScope)"
    }
    else {
      # Use AppId directly — always unambiguous in EXO
      New-ManagementRoleAssignment `
      -Name                $raName `
      -Role                $roleSpec.Role `
      -App                 $script:app.AppId `
      -CustomResourceScope $ScopeName `
      -ErrorAction         Stop | Out-Null

      # Verify scope was actually persisted
      $verify = Get-ManagementRoleAssignment -Identity $raName -ErrorAction Stop
      if ([string]::IsNullOrWhiteSpace($verify.CustomResourceScope)) {
        throw ("Role assignment '$raName' was created but CustomResourceScope " +
          "is empty — the assignment is tenant-wide. " +
        "Check that '$ScopeName' is a valid management scope name.")
      }
      if ($verify.CustomResourceScope -ne $ScopeName) {
        throw ("Role assignment '$raName' has unexpected scope " +
        "'$($verify.CustomResourceScope)' — expected '$ScopeName'.")
      }

      Write-Info "Assigned: $($roleSpec.Role) → $ScopeName"
      Write-Info "Verified scope on new assignment: $($verify.CustomResourceScope)"
    }
  }
}

#endregion

#region ── Step 12: Verify Role Assignments ───────────────────────────────────

Write-Step "Verifying role assignments against scoped mailboxes"

Invoke-Step -Name "Role assignment verification" -ContinueOnError -Action {
  $allPassed = $true

  foreach ($mbx in $script:validatedMailboxes) {
    $results = Test-ServicePrincipalAuthorization `
    -Identity $AppName `
    -Resource $mbx.PrimarySmtpAddress `
    -ErrorAction Stop

    foreach ($r in $results) {
      $status = if ($r.InScope) { "✔" } else { "✘" }
      $color  = if ($r.InScope) { "Green" } else { "Red" }
      Write-Host ("  {0}  {1,-35} {2,-25} InScope={3}" -f `
      $status, $mbx.PrimarySmtpAddress, $r.RoleName, $r.InScope) `
      -ForegroundColor $color

      if (-not $r.InScope) { $allPassed = $false }
    }
  }

  if (-not $allPassed) {
    Write-Warn "Some mailboxes did not pass scope verification."
    Write-Warn "This may resolve within 30–120 minutes due to permission caching."
    Write-Warn "Run Test-ScopedMailboxApp.ps1 after the propagation window."
  }
  else {
    Write-Info "All mailboxes verified in-scope."
  }
}

#endregion

#region ── Step 13: Write Output File ─────────────────────────────────────────

Write-Step "Writing configuration summary"

Invoke-Step -Name "Output file" -Action {
    $outputFile = Join-Path $OutputPath "$AppName-config.json"
    $script:config | ConvertTo-Json -Depth 5 | Set-Content -Path $outputFile -Encoding UTF8
    Write-Info "Configuration written to: $outputFile"
    Write-Warn "This file contains sensitive credentials — store it securely."
}

#endregion

#region ── Final Summary ───────────────────────────────────────────────────────

Write-Host ""
Write-Host ("═" * 70) -ForegroundColor DarkCyan
Write-Host "  CONFIGURATION COMPLETE" -ForegroundColor Cyan
Write-Host ("═" * 70) -ForegroundColor DarkCyan
Write-Host "  App Name       : $AppName"
Write-Host "  Client ID      : $($script:config.AppId)"
Write-Host "  Tenant ID      : $($script:config.TenantId)"
Write-Host "  Scope          : $ScopeName  ($ScopeFilter)"
Write-Host "  MESG           : $MESGName"
Write-Host "  Mailboxes      : $($Mailboxes -join ', ')"
Write-Host "  Config File    : $(Join-Path $OutputPath "$AppName-config.json")"
Write-Host ""
Write-Host "  Next step: Run Test-ScopedMailboxApp.ps1 after 30–120 minutes"
Write-Host "             to allow permission propagation."
Write-Host ("═" * 70) -ForegroundColor DarkCyan

Write-SummaryAndExit -ExitCode 0

#endregion
