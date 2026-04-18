#Requires -Version 5.1
<#
    .SYNOPSIS
    Creates an Entra ID app registration for app-only SMTP OAuth2 sending
    via Exchange Online, using the SMTP.SendAsApp application permission.

    .DESCRIPTION
    Implements the client credentials flow for SMTP AUTH as documented at:
    https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth

    The correct flow requires:
      1. SMTP.SendAsApp application permission on Office 365 Exchange Online
         (AppId 00000002-0000-0ff1-ce00-000000000000) — NOT on Microsoft Graph
         and NOT the delegated SMTP.Send permission.
      2. Admin consent granted for the tenant.
      3. The app's service principal registered in Exchange Online via
         New-ServicePrincipal using the Enterprise Application Object ID
         (not the App Registration Object ID).
      4. Mailbox access granted per-mailbox via Add-MailboxPermission.
      5. Token acquired using client credentials flow with scope
         https://outlook.office365.com/.default
      6. XOAUTH2 string built with user= set to the sending mailbox address
         and auth=Bearer <token>.

    This is fully app-only — no interactive user sign-in is required at any
    point, shared mailboxes are supported, and mailbox access is explicitly
    scoped per-mailbox via Add-MailboxPermission.

    This script:
      1. Creates an Entra ID app registration (single-tenant)
      2. Adds the SMTP.SendAsApp application permission on the Office 365
         Exchange Online resource
      3. Grants admin consent
      4. Creates a client secret or self-signed certificate
      5. Creates the service principal in Exchange Online
      6. Grants the service principal FullAccess to each specified mailbox
         via Add-MailboxPermission
      7. Enables SMTP AUTH per-mailbox (SmtpClientAuthenticationDisabled)
      8. Checks the org-level SMTP AUTH setting
      9. Writes a config JSON for use by Test-SmtpOAuthApp.ps1

    .PARAMETER AppName
    Display name for the app registration.

    .PARAMETER Mailboxes
    SMTP addresses of the mailboxes the app will send as.
    Shared mailboxes are supported.

    .PARAMETER UseCertificate
    Creates a self-signed certificate credential instead of a client secret.

    .PARAMETER ExpiryMonths
    Months until the certificate or client secret expires. Defaults to 12.

    .PARAMETER OutputPath
    Directory for the config JSON and certificate file.
    Defaults to the current directory.

    .EXAMPLE
    .\New-SmtpOAuthApp.ps1 `
        -AppName   "SmtpMailer-HR" `
        -Mailboxes @("hr@contoso.com", "notifications@contoso.com")

    .EXAMPLE
    .\New-SmtpOAuthApp.ps1 `
        -AppName        "SmtpMailer-HR" `
        -Mailboxes      @("hr@contoso.com") `
        -UseCertificate

    .NOTES
    Requires:
      - Microsoft.Graph.Applications module
      - ExchangeOnlineManagement module
      - Entra ID Application Administrator or higher
      - Exchange Online administrator
      - Network access to login.microsoftonline.com (for admin consent)
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
    [ValidateRange(1, 24)]
    [int]$ExpiryMonths = 12,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Get-Location).Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Compatible module versions ───────────────────────────────────────────────
# ExchangeOnlineManagement 3.8+ and Microsoft.Graph 2.25+ cannot coexist in
# the same PowerShell session due to a WAM/MSAL broker DLL conflict.
# These versions are the last known-good combination where both load cleanly.
$script:RequiredExoVersion   = "3.7.0"
$script:RequiredGraphVersion = "2.24.0"

# ── Fixed identifiers ───────────────────────────────────────────────────────
# Office 365 Exchange Online first-party service principal.
# This AppId is the same in every tenant — do not change it.
$script:ExoResourceAppId  = "00000002-0000-0ff1-ce00-000000000000"
$script:ExoResourceUrl    = "https://outlook.office365.com"

# The SMTP.SendAsApp permission ID on the EXO resource.
# Resolved dynamically from the EXO service principal below to guard
# against future Microsoft changes, but this is the current known value.
$script:SmtpSendAsAppId   = $null

#region ── Helpers ─────────────────────────────────────────────────────────────

$script:StepNumber  = 0
$script:StepsFailed = [System.Collections.Generic.List[string]]::new()

$script:app                = $null
$script:sp                 = $null
$script:exoSp              = $null
$script:tenantId           = $null
$script:validatedMailboxes = [System.Collections.Generic.List[object]]::new()
$script:config             = $null

function Write-Step {
    param([string]$Message)
    $script:StepNumber++
    Write-Host ""
    Write-Host ("─" * 70) -ForegroundColor DarkGray
    Write-Host "  STEP $($script:StepNumber) │ $Message" -ForegroundColor Cyan
    Write-Host ("─" * 70) -ForegroundColor DarkGray
}

function Write-Ok   { param([string]$m) Write-Host "  ✔  $m" -ForegroundColor Green    }
function Write-Info { param([string]$m) Write-Host "  ℹ  $m" -ForegroundColor DarkCyan }
function Write-Warn { param([string]$m) Write-Host "  ⚠  $m" -ForegroundColor Yellow   }

function Assert-NotNull {
    param([object]$Value, [string]$Label)
    if ($null -eq $Value -or
        ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) {
        throw "$Label was null or empty."
    }
}

function Get-ValueOrDefault {
    param([object]$Value, [object]$Default)
    if ($null -ne $Value -and
        -not ($Value -is [string] -and [string]::IsNullOrEmpty($Value))) {
        return $Value
    }
    return $Default
}

function Invoke-Step {
    param(
        [string]$Name,
        [scriptblock]$Action,
        [switch]$ContinueOnError
    )
    try {
        & $Action
        Write-Ok "$Name completed."
    }
    catch {
        $script:StepsFailed.Add($Name)
        Write-Host "  ✘  $Name FAILED: $_" -ForegroundColor Red
        if (-not $ContinueOnError) {
            Write-Host "  ► Aborting." -ForegroundColor Red
            Write-FinalSummary -ExitCode 1
        }
        else {
            Write-Warn "Continuing despite error (ContinueOnError set)."
        }
    }
}

function Write-FinalSummary {
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
Write-Host "  New-SmtpOAuthApp  │  App-Only SMTP OAuth2 Setup" -ForegroundColor Cyan
Write-Host ("═" * 70) -ForegroundColor DarkCyan
Write-Host "  App Name    : $AppName"
Write-Host "  Mailboxes   : $($Mailboxes -join ', ')"
Write-Host "  Credential  : $(if ($UseCertificate) { 'Certificate' } else { 'Client Secret' })"
Write-Host "  Expiry      : $ExpiryMonths month(s)"
Write-Host "  Output Path : $OutputPath"
Write-Host ("═" * 70) -ForegroundColor DarkCyan

#endregion

#region ── Config skeleton ─────────────────────────────────────────────────────

$script:config = [ordered]@{
    AppName          = $AppName
    Mailboxes        = $Mailboxes
    TenantId         = $null
    AppId            = $null
    AppObjectId      = $null
    SpObjectId       = $null
    ExoSpIdentity    = $null
    CredentialType   = if ($UseCertificate) { "Certificate" } else { "ClientSecret" }
    ExpiryMonths     = $ExpiryMonths
    ClientSecret     = $null
    CertThumbprint   = $null
    TokenResource    = $script:ExoResourceUrl
    TokenScope       = "$($script:ExoResourceUrl)/.default"
    SmtpHost         = "smtp.office365.com"
    SmtpPort         = 587
    ConsentUrl       = $null
    SmtpAuthResults  = @{}
    MailboxPermissions = @{}
    CreatedAt        = (Get-Date -Format "o")
    Notes            = @(
        "Permission: SMTP.SendAsApp application permission on Office 365 Exchange Online.",
        "Flow: Client credentials — no interactive sign-in required.",
        "Token scope: https://outlook.office365.com/.default",
        "XOAUTH2 user= field must be set to the sending mailbox address.",
        "Shared mailboxes are supported.",
        "Mailbox access is controlled per-mailbox via Add-MailboxPermission.",
        "Reference: https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth"
    )
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


Invoke-Step -Name "Output path" -Action {
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-Info "Created: $OutputPath"
    }
}

#endregion

#region ── Step 1: Connect to Microsoft Graph ──────────────────────────────────

Write-Step "Connecting to Microsoft Graph"

Invoke-Step -Name "Module check" -Action {
    # Use specific sub-modules only — never the Microsoft.Graph meta-module.
    # The meta-module imports 38+ sub-modules and exceeds the 4096-function
    # limit in Windows PowerShell 5.1, preventing Microsoft.Graph.Authentication
    # from loading correctly.
    $required = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Applications",
        "ExchangeOnlineManagement"
    )
    foreach ($mod in $required) {
        if (-not (Get-Module -ListAvailable -Name $mod)) {
            throw ("Module '$mod' is not installed. " +
                   "Run: Install-Module $mod -Scope CurrentUser")
        }
        Write-Info "Module found: $mod"
    }
}

#endregion

#region ── Step 2: Connect to Microsoft Graph ──────────────────────────────────

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

    $script:tenantId        = $ctx.TenantId
    $script:config.TenantId = $ctx.TenantId
    Write-Info "Tenant : $($ctx.TenantId)"
    Write-Info "Account: $($ctx.Account)"
}

#endregion

#region ── Step 2a: Connect to Exchange Online ──────────────────────────────────

Write-Step "Connecting to Exchange Online"

Invoke-Step -Name "EXO connection" -Action {
    Import-Module ExchangeOnlineManagement `
        -RequiredVersion $script:RequiredExoVersion `
        -ErrorAction Stop

    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Info "Already connected to Exchange Online."
    }
    catch {
        Connect-ExchangeOnline `
            -Organization  $script:tenantId `
            -ShowBanner:$false `
            -ErrorAction   Stop

        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Info "Connected to Exchange Online (tenant: $script:tenantId)."
    }
}


#endregion

#region ── Step 3: Validate Mailboxes ─────────────────────────────────────────

Write-Step "Validating mailboxes"

Invoke-Step -Name "Mailbox validation" -Action {
    foreach ($addr in $Mailboxes) {
        $mbx = Get-Mailbox -Identity $addr -ErrorAction Stop
        $script:validatedMailboxes.Add($mbx)
        Write-Info "Validated: $addr  (Type: $($mbx.RecipientTypeDetails))"
    }
    Write-Info "$($script:validatedMailboxes.Count) mailbox(es) validated."
}

#endregion

#region ── Step 4: Check Org-Level SMTP AUTH ──────────────────────────────────

Write-Step "Checking organisation-level SMTP AUTH setting"

Invoke-Step -Name "Org SMTP AUTH check" -ContinueOnError -Action {
    $transport = Get-TransportConfig -ErrorAction Stop
    $val       = $transport.SmtpClientAuthenticationDisabled

    if ($val -eq $true) {
        Write-Info ("Organisation-level SMTP AUTH is globally disabled, " +
                    "but per-mailbox enablement (SmtpClientAuthenticationDisabled = `$false) " +
                    "overrides this for individual mailboxes. No org-level change required.")
        $script:config.OrgSmtpAuthDisabled = $true
    }
    else {
        Write-Info ("Org-level SMTP AUTH setting: $val.")
        $script:config.OrgSmtpAuthDisabled = $false
    }
}

#endregion

#region ── Step 5: Create App Registration ─────────────────────────────────────

Write-Step "Creating Entra ID app registration"

Invoke-Step -Name "App registration" -Action {
    $existing = Get-MgApplication `
        -Filter      "displayName eq '$AppName'" `
        -ErrorAction SilentlyContinue |
        Select-Object -First 1

    if ($existing) {
        Write-Warn "App '$AppName' already exists (AppId: $($existing.AppId)). Using existing."
        $script:app = $existing
    }
    else {
        New-MgApplication `
            -DisplayName    $AppName `
            -SignInAudience "AzureADMyOrg" | Out-Null

        Write-Info "App created — waiting for replication..."
        Start-Sleep -Seconds 5

        $script:app = Get-MgApplication `
            -Filter      "displayName eq '$AppName'" `
            -ErrorAction Stop |
            Select-Object -First 1
    }

    Assert-NotNull -Value $script:app       -Label "App object"
    Assert-NotNull -Value $script:app.AppId -Label "App AppId"
    Assert-NotNull -Value $script:app.Id    -Label "App Object ID"

    $script:config.AppId       = $script:app.AppId
    $script:config.AppObjectId = $script:app.Id

    Write-Info "App ID (Client ID) : $($script:app.AppId)"
    Write-Info "App Object ID      : $($script:app.Id)"
}

#endregion

#region ── Step 6: Create Entra Service Principal ──────────────────────────────

Write-Step "Creating Entra ID service principal"

Invoke-Step -Name "Entra service principal" -Action {
    $existing = Get-MgServicePrincipal `
        -Filter      "appId eq '$($script:app.AppId)'" `
        -ErrorAction SilentlyContinue |
        Select-Object -First 1

    if ($existing) {
        Write-Warn "Service principal already exists. Using existing."
        $script:sp = $existing
    }
    else {
        New-MgServicePrincipal -AppId $script:app.AppId -ErrorAction Stop | Out-Null
        Start-Sleep -Seconds 5

        $script:sp = Get-MgServicePrincipal `
            -Filter      "appId eq '$($script:app.AppId)'" `
            -ErrorAction Stop |
            Select-Object -First 1
    }

    Assert-NotNull -Value $script:sp    -Label "Service principal"
    Assert-NotNull -Value $script:sp.Id -Label "SP Object ID"

    $script:config.SpObjectId = $script:sp.Id
    Write-Info "Entra SP Object ID : $($script:sp.Id)"
    Write-Info ("Note: The EXO New-ServicePrincipal step uses this " +
                "Object ID, not the App Object ID.")
}

#endregion

#region ── Step 7: Add SMTP.SendAsApp Application Permission ──────────────────
#
# SMTP.SendAsApp is an APPLICATION permission on the Office 365 Exchange
# Online resource (AppId 00000002-0000-0ff1-ce00-000000000000).
#
# This is different from:
#   - SMTP.Send  (delegated permission — requires user sign-in, does not
#                 work with client credentials flow for SMTP AUTH)
#   - Microsoft Graph SMTP permissions (wrong resource entirely)
#
# The permission ID is resolved dynamically from the EXO service principal
# to avoid hardcoding a value that Microsoft could change.

Write-Step "Adding SMTP.SendAsApp application permission (EXO resource)"

Invoke-Step -Name "SMTP.SendAsApp permission" -Action {
    # Locate the Office 365 Exchange Online first-party service principal
    $exoResource = Get-MgServicePrincipal `
        -Filter      "appId eq '$script:ExoResourceAppId'" `
        -ErrorAction Stop |
        Select-Object -First 1

    if (-not $exoResource) {
        throw ("Office 365 Exchange Online service principal not found. " +
               "Ensure Exchange Online is provisioned in this tenant.")
    }

    Write-Info "EXO resource SP found: $($exoResource.DisplayName)"

    # Find SMTP.SendAsApp in the app roles (application permissions)
    $smtpRole = $exoResource.AppRoles |
        Where-Object { $_.Value -eq "SMTP.SendAsApp" } |
        Select-Object -First 1

    if (-not $smtpRole) {
        # List available app roles to aid diagnosis
        $available = ($exoResource.AppRoles | Select-Object -ExpandProperty Value) -join ", "
        throw ("SMTP.SendAsApp app role not found on EXO resource. " +
               "Available roles: $available")
    }

    $script:SmtpSendAsAppId = $smtpRole.Id
    Write-Info "SMTP.SendAsApp role ID: $($smtpRole.Id)"

    # Check if already assigned
    $existingAssignment = Get-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $script:sp.Id `
        -ErrorAction        SilentlyContinue |
        Where-Object {
            $_.AppRoleId   -eq $smtpRole.Id -and
            $_.ResourceId  -eq $exoResource.Id
        }

    if ($existingAssignment) {
        Write-Warn "SMTP.SendAsApp already assigned. Skipping."
        return
    }

    # Assign the app role (this is equivalent to adding the permission AND
    # granting admin consent in a single operation when done via PowerShell)
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $script:sp.Id `
        -PrincipalId        $script:sp.Id `
        -ResourceId         $exoResource.Id `
        -AppRoleId          $smtpRole.Id `
        -ErrorAction        Stop | Out-Null

    Write-Info "SMTP.SendAsApp application permission assigned and consented."
    Write-Info ("This grants the app permission to send as any mailbox " +
                "that is explicitly authorised via Add-MailboxPermission.")
}

#endregion

#region ── Step 8: Create Credential ──────────────────────────────────────────

Write-Step "Creating app credential"

Invoke-Step -Name "Credential creation" -Action {
    if ($UseCertificate) {
        $cert = New-SelfSignedCertificate `
            -Subject           "CN=$AppName" `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -KeyExportPolicy   Exportable `
            -KeySpec           Signature `
            -KeyLength         2048 `
            -HashAlgorithm     SHA256 `
            -NotAfter          (Get-Date).AddMonths($ExpiryMonths)

        $certPath  = Join-Path $OutputPath "$AppName.cer"
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
                    EndDateTime = (Get-Date).AddMonths($ExpiryMonths)
                }
            }

        Assert-NotNull -Value $secretResult.SecretText -Label "Client secret"
        $script:config.ClientSecret = $secretResult.SecretText

        $expiry = (Get-Date).AddMonths($ExpiryMonths).ToString("yyyy-MM-dd")
        Write-Info "Client secret created (expires: $((Get-Date).AddMonths($ExpiryMonths).ToString('yyyy-MM-dd')))"
        Write-Warn "Client secret is in the output file — store it securely."
    }
}

#endregion

#region ── Step 9: Register Service Principal in Exchange Online ───────────────
#
# This step is mandatory. Without it, Exchange Online cannot resolve which
# app is authenticating when it receives the client credentials token.
# The OBJECT_ID used here must be the Enterprise Application Object ID
# (from the service principal), NOT the App Registration Object ID.
# Using the wrong ID causes authentication failures.

Write-Step "Registering service principal in Exchange Online"

Invoke-Step -Name "EXO service principal registration" -Action {
    Assert-NotNull -Value $script:app.AppId -Label "App AppId"
    Assert-NotNull -Value $script:sp.Id     -Label "Entra SP Object ID"

    # Check if already registered in EXO
    $existing = Get-ServicePrincipal -ErrorAction SilentlyContinue |
        Where-Object { $_.AppId -eq $script:app.AppId } |
        Select-Object -First 1

    if ($existing) {
        Write-Warn "EXO service principal already registered. Using existing."
        $script:exoSp = $existing
    }
    else {
        # Retry loop to handle Entra replication lag
        $maxAttempts = 5
        $retryWait   = 15
        $attempt     = 0
        $lastError   = $null

        while ($attempt -lt $maxAttempts -and -not $script:exoSp) {
            $attempt++
            try {
                Write-Info ("Registering EXO SP " +
                            "(attempt $attempt of $maxAttempts)...")

                # IMPORTANT: -ObjectId must be the Entra SERVICE PRINCIPAL
                # Object ID, not the App Registration Object ID.
                $script:exoSp = New-ServicePrincipal `
                    -AppId       $script:app.AppId `
                    -ObjectId    $script:sp.Id `
                    -DisplayName $AppName `
                    -ErrorAction Stop
            }
            catch {
                $lastError = $_
                if ($attempt -lt $maxAttempts) {
                    Write-Warn "Attempt $attempt failed: $lastError"
                    Write-Info "Waiting ${retryWait}s before retry..."
                    Start-Sleep -Seconds $retryWait
                }
            }
        }

        if (-not $script:exoSp) {
            throw ("Failed to register EXO service principal after " +
                   "$maxAttempts attempts. Last error: $lastError")
        }
    }

    Assert-NotNull -Value $script:exoSp          -Label "EXO service principal"
    Assert-NotNull -Value $script:exoSp.Identity -Label "EXO SP Identity"

    $script:config.ExoSpIdentity = $script:exoSp.Identity
    Write-Info "EXO SP Identity: $($script:exoSp.Identity)"
}

#endregion

#region ── Step 10: Grant Mailbox Permissions ──────────────────────────────────
#
# Add-MailboxPermission grants the service principal FullAccess to each
# mailbox. This is the per-mailbox scoping mechanism for SMTP.SendAsApp —
# the app can only send as mailboxes that have been explicitly granted here.
#
# Note: For SMTP sending specifically, FullAccess is required even though
# the app is only sending (not reading). This is the documented requirement
# for the SMTP.SendAsApp flow. If SendAs permission is also required
# (e.g. for the MAIL FROM address), Add-RecipientPermission -AccessRights
# SendAs should also be run — this is noted in the Microsoft documentation.

Write-Step "Granting mailbox permissions to service principal"

Invoke-Step -Name "Mailbox permissions" -Action {
    Assert-NotNull -Value $script:exoSp.Identity -Label "EXO SP Identity"

    foreach ($mbx in $script:validatedMailboxes) {
        $addr = $mbx.PrimarySmtpAddress

        try {
            # FullAccess — required for SMTP.SendAsApp client credentials flow
            Add-MailboxPermission `
                -Identity     $addr `
                -User         $script:exoSp.Identity `
                -AccessRights FullAccess `
                -InheritanceType All `
                -Confirm:$false `
                -ErrorAction  Stop | Out-Null

            Write-Info "FullAccess granted: $addr"

            # SendAs — required if the app sets MAIL FROM to the mailbox address
            Add-RecipientPermission `
                -Identity    $addr `
                -Trustee     $script:exoSp.Identity `
                -AccessRights SendAs `
                -Confirm:$false `
                -ErrorAction Stop | Out-Null

            Write-Info "SendAs granted: $addr"
            $script:config.MailboxPermissions[$addr] = "FullAccess, SendAs"
        }
        catch {
            if ($_ -match "already exists|duplicate") {
                Write-Warn "Permission already exists for $addr — skipping."
                $script:config.MailboxPermissions[$addr] = "Already granted"
            }
            else {
                throw $_
            }
        }
    }
}

#endregion

#region ── Step 11: Enable SMTP AUTH Per-Mailbox ──────────────────────────────

Write-Step "Enabling SMTP AUTH on mailboxes"

Invoke-Step -Name "Per-mailbox SMTP AUTH" -Action {
    foreach ($mbx in $script:validatedMailboxes) {
        $addr       = $mbx.PrimarySmtpAddress
        $casMailbox = Get-CASMailbox -Identity $addr -ErrorAction Stop
        $current    = $casMailbox.SmtpClientAuthenticationDisabled

        $statusText = switch ($current) {
            $true  { "explicitly disabled" }
            $false { "already explicitly enabled" }
            $null  { "inheriting org setting" }
        }

        Write-Info "Current SMTP AUTH for $addr : $statusText"

        if ($current -ne $false) {
            Set-CASMailbox `
                -Identity                         $addr `
                -SmtpClientAuthenticationDisabled $false `
                -ErrorAction                      Stop

            $verify = Get-CASMailbox -Identity $addr -ErrorAction Stop
            if ($verify.SmtpClientAuthenticationDisabled -ne $false) {
                throw ("Failed to enable SMTP AUTH for '$addr'. " +
                       "Value after Set-CASMailbox: " +
                       "$($verify.SmtpClientAuthenticationDisabled)")
            }

            Write-Info "Enabled SMTP AUTH for $addr (was: $statusText)"
            $script:config.SmtpAuthResults[$addr] = "Enabled (was: $statusText)"
        }
        else {
            Write-Info "No change needed for $addr."
            $script:config.SmtpAuthResults[$addr] = "Already enabled — no change"
        }
    }
}

#endregion

#region ── Step 12: Build Consent URL ─────────────────────────────────────────

Write-Step "Building admin consent URL"

Invoke-Step -Name "Consent URL" -Action {
    # The scope for SMTP.SendAsApp consent uses the EXO resource URL.
    # Reference from Microsoft docs:
    # "scope query parameter should be https://outlook.office365.com/.default
    #  only for SMTP"
    $consentUrl = ("https://login.microsoftonline.com/$($script:config.TenantId)" +
                   "/v2.0/adminconsent" +
                   "?client_id=$($script:config.AppId)" +
                   "&redirect_uri=https://localhost" +
                   "&scope=$($script:ExoResourceUrl)/.default")

    $script:config.ConsentUrl = $consentUrl
    Write-Info "Consent URL:"
    Write-Info "  $consentUrl"
    Write-Info ("Note: Admin consent was already granted via " +
                "New-MgServicePrincipalAppRoleAssignment in Step 7. " +
                "This URL is provided for reference and verification.")
}

#endregion

#region ── Step 13: Write Output File ─────────────────────────────────────────

Write-Step "Writing configuration file"

Invoke-Step -Name "Config file" -Action {
    $outputFile = Join-Path $OutputPath "$AppName-smtp-config.json"
    $script:config | ConvertTo-Json -Depth 5 |
        Set-Content -Path $outputFile -Encoding UTF8
    Write-Info "Written to: $outputFile"
    Write-Warn "File contains credentials — store securely."
}

#endregion

#region ── Final Banner ────────────────────────────────────────────────────────

Write-Host ""
Write-Host ("═" * 70) -ForegroundColor DarkCyan
Write-Host "  SMTP OAUTH2 APP SETUP COMPLETE" -ForegroundColor Cyan
Write-Host ("═" * 70) -ForegroundColor DarkCyan
Write-Host "  App Name       : $AppName"
Write-Host "  Client ID      : $($script:config.AppId)"
Write-Host "  Tenant ID      : $($script:config.TenantId)"
Write-Host "  EXO SP Identity: $($script:config.ExoSpIdentity)"
Write-Host "  Token Resource : $($script:config.TokenResource)"
Write-Host "  Token Scope    : $($script:config.TokenScope)"
Write-Host "  SMTP Host      : $($script:config.SmtpHost):$($script:config.SmtpPort)"
Write-Host ""
Write-Host "  Mailboxes:"
foreach ($addr in $script:config.MailboxPermissions.Keys) {
    Write-Host ("    {0,-45} Permissions: {1}" -f `
        $addr, $script:config.MailboxPermissions[$addr])
}
Write-Host ""
Write-Host "  SMTP AUTH status:"
foreach ($addr in $script:config.SmtpAuthResults.Keys) {
    Write-Host ("    {0,-45} {1}" -f $addr, $script:config.SmtpAuthResults[$addr])
}
Write-Host ""
Write-Host "  Next step: Run Test-SmtpOAuthApp.ps1 to validate."
Write-Host "  Config file: $(Join-Path $OutputPath "$AppName-smtp-config.json")"
Write-Host ("═" * 70) -ForegroundColor DarkCyan

Write-FinalSummary -ExitCode 0

#endregion
