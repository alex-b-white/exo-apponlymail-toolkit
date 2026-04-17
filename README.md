# Exchange Online App-Only Mail Access — PowerShell Toolkit

Automate the creation and validation of **scoped, app-only** Exchange Online mailbox access using Entra ID app registrations. Covers both the **Microsoft Graph API** (scoped RBAC) and **SMTP OAuth2** (`SMTP.SendAsApp`) flows.

---

## Scripts

| Script | Purpose |
|---|---|
| [`New-ScopedMailboxApp.ps1`](#new-scopedmailboxappps1) | Creates an Entra ID app with Graph API access scoped to specific mailboxes via EXO management scopes |
| [`Test-ScopedMailboxApp.ps1`](#test-scopedmailboxappps1) | Validates RBAC configuration and runs live Graph API read/send tests |
| [`New-SmtpOAuthApp.ps1`](#new-smtpoauthappps1) | Creates an Entra ID app with `SMTP.SendAsApp` permission for app-only SMTP OAuth2 sending |
| [`Test-SmtpOAuthApp.ps1`](#test-smtpoauthappps1) | Validates the full SMTP OAuth2 flow including token claims, STARTTLS, XOAUTH2, and message delivery |

---

## Prerequisites

### PowerShell Version

Requires **PowerShell 5.1** or later.

### Module Installation

> **⚠️ Important — Module Compatibility**
>
> `ExchangeOnlineManagement` 3.8+ and `Microsoft.Graph` 2.25+ **cannot coexist in the same PowerShell session** due to a WAM/MSAL broker DLL conflict ([msgraph-sdk-powershell #3576](https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3576)). Install the last known-good versions below. This is a **side-by-side install** — your existing module versions are not removed.

```powershell
# Install the last compatible versions not affected by Graph SDK PowerShell issue #3576
# (https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3576)
# Side-by-side install — does not remove your existing versions
Install-Module ExchangeOnlineManagement `
    -RequiredVersion 3.7.0 `
    -Scope CurrentUser `
    -Force `
    -AllowClobber

Install-Module Microsoft.Graph.Authentication `
    -RequiredVersion 2.24.0 `
    -Scope CurrentUser `
    -Force `
    -AllowClobber

Install-Module Microsoft.Graph.Applications `
    -RequiredVersion 2.24.0 `
    -Scope CurrentUser `
    -Force `
    -AllowClobber

# Verify all three installed correctly
Get-InstalledModule ExchangeOnlineManagement,
                    Microsoft.Graph.Authentication,
                    Microsoft.Graph.Applications |
    Select-Object Name, Version |
    Sort-Object Name
```

### Required Permissions

| Role | Required For |
|---|---|
| Entra ID Application Administrator (or higher) | Creating app registrations and service principals |
| Exchange Online Administrator | Management scopes, role assignments, mailbox permissions |
| Network access to `smtp.office365.com:587` | SMTP OAuth2 tests only |

---

## Script Reference

### `New-ScopedMailboxApp.ps1`

Creates an Entra ID app registration with **Graph API** mailbox access scoped to a specific list of mailboxes using Exchange Online management scopes and custom attributes.

**What it does:**

1. Creates an Entra ID app registration and service principal
2. Creates a client secret or self-signed certificate
3. Creates a mail-enabled security group (MESG) and adds the specified mailboxes
4. Stamps a custom attribute on each mailbox for scope filtering
5. Creates an Exchange Online management scope
6. Registers the EXO service principal
7. Assigns scoped `Application Mail.Read` and `Application Mail.Send` roles
8. Writes a `<AppName>-config.json` file for use with the test script

**Usage:**

```powershell
# Client secret (default)
.\New-ScopedMailboxApp.ps1 `
    -AppName   "MailApp-HR" `
    -Mailboxes @("hr@contoso.com", "payroll@contoso.com")

# Certificate credential
.\New-ScopedMailboxApp.ps1 `
    -AppName         "MailApp-Finance" `
    -Mailboxes       @("finance@contoso.com") `
    -UseCertificate `
    -CustomAttribute "CustomAttribute2"
```

**Parameters:**

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-AppName` | ✅ | — | Display name for the app and associated resources |
| `-Mailboxes` | ✅ | — | Array of mailbox SMTP addresses to scope access to |
| `-UseCertificate` | ❌ | `$false` | Create a self-signed certificate instead of a client secret |
| `-CustomAttribute` | ❌ | `CustomAttribute15` | Mailbox custom attribute (1–15) used for scope filtering |
| `-SecretExpiryMonths` | ❌ | `6` | Client secret validity period in months |
| `-OutputPath` | ❌ | Current directory | Directory for the config JSON and certificate file |

---

### `Test-ScopedMailboxApp.ps1`

Validates the configuration created by `New-ScopedMailboxApp.ps1` and runs live Graph API tests.

**What it tests:**

| Category | Tests |
|---|---|
| **RBAC** | Management scope exists with correct filter; role assignments exist and are correctly scoped; EXO service principal registered; MESG membership; custom attribute stamps; `Test-ServicePrincipalAuthorization` results |
| **Graph API** | Token acquisition; read mail from each in-scope mailbox (expect 200); send mail from each in-scope mailbox (expect 202); read from out-of-scope mailbox (expect 403/404); send from out-of-scope mailbox (expect 403/404) |

> **Note on out-of-scope denial testing:** `Test-ServicePrincipalAuthorization` results for out-of-scope mailboxes are reported as **informational only**. The cmdlet does not reliably report denial for non-exclusive management scopes. The Graph API HTTP response (403/404) is the authoritative denial check.

**Usage:**

```powershell
# Full test with out-of-scope denial check
.\Test-ScopedMailboxApp.ps1 `
    -AppName           "MailApp-HR" `
    -Mailboxes         @("hr@contoso.com", "payroll@contoso.com") `
    -ConfigFile        ".\MailApp-HR-config.json" `
    -OutOfScopeMailbox "finance@contoso.com"

# RBAC checks only — skip live API calls
.\Test-ScopedMailboxApp.ps1 `
    -AppName      "MailApp-HR" `
    -Mailboxes    @("hr@contoso.com") `
    -ConfigFile   ".\MailApp-HR-config.json" `
    -SkipApiTests
```

**Parameters:**

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-AppName` | ✅ | — | Display name of the app registration to test |
| `-Mailboxes` | ✅ | — | In-scope mailbox addresses to test |
| `-ConfigFile` | ❌ | — | Path to JSON config from `New-ScopedMailboxApp.ps1`; auto-populates credentials |
| `-TenantId` | ❌* | — | Entra ID tenant ID (*required if no `-ConfigFile`) |
| `-AppId` | ❌* | — | Application client ID (*required if no `-ConfigFile`) |
| `-ClientSecret` | ❌* | — | Client secret (*required if no `-ConfigFile` and not using cert) |
| `-CertThumbprint` | ❌* | — | Certificate thumbprint (*required if no `-ConfigFile` and using cert) |
| `-OutOfScopeMailbox` | ❌ | — | A real EXO mailbox **not** in scope — used to verify access is denied |
| `-SkipApiTests` | ❌ | `$false` | Run RBAC checks only; skip live Graph API calls |
| `-TestRecipient` | ❌ | Sending mailbox | Recipient for test messages (defaults to loop-back) |

**Exit codes:** `0` = all tests passed, `1` = one or more tests failed.

---

### `New-SmtpOAuthApp.ps1`

Creates an Entra ID app registration for **app-only SMTP OAuth2 sending** via Exchange Online using the `SMTP.SendAsApp` client credentials flow.

**What it does:**

1. Creates an Entra ID app registration and service principal
2. Adds the `SMTP.SendAsApp` **application** permission on the Office 365 Exchange Online resource (`00000002-0000-0ff1-ce00-000000000000`) — **not** Microsoft Graph
3. Grants admin consent via `New-MgServicePrincipalAppRoleAssignment`
4. Creates a client secret or self-signed certificate
5. Registers the service principal in Exchange Online
6. Grants `FullAccess` and `SendAs` per-mailbox via `Add-MailboxPermission` / `Add-RecipientPermission`
7. Enables SMTP AUTH per-mailbox (`SmtpClientAuthenticationDisabled = $false`)
8. Checks the org-level SMTP AUTH setting
9. Writes a `<AppName>-smtp-config.json` file for use with the test script

> **⚠️ Common mistakes this script avoids:**
> - Using `SMTP.Send` (delegated) instead of `SMTP.SendAsApp` (application)
> - Adding the permission on Microsoft Graph instead of the EXO resource
> - Using the App Registration Object ID instead of the Enterprise Application (service principal) Object ID when calling `New-ServicePrincipal`
> - Using `https://graph.microsoft.com/.default` as the token scope instead of `https://outlook.office365.com/.default`

**Usage:**

```powershell
# Client secret
.\New-SmtpOAuthApp.ps1 `
    -AppName   "SmtpMailer-HR" `
    -Mailboxes @("hr@contoso.com", "notifications@contoso.com")

# Certificate credential
.\New-SmtpOAuthApp.ps1 `
    -AppName        "SmtpMailer-HR" `
    -Mailboxes      @("hr@contoso.com") `
    -UseCertificate
```

**Parameters:**

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-AppName` | ✅ | — | Display name for the app registration |
| `-Mailboxes` | ✅ | — | SMTP addresses of mailboxes the app will send as (shared mailboxes supported) |
| `-UseCertificate` | ❌ | `$false` | Create a self-signed certificate instead of a client secret |
| `-SecretExpiryMonths` | ❌ | `12` | Client secret validity period in months |
| `-OutputPath` | ❌ | Current directory | Directory for the config JSON and certificate file |

---

### `Test-SmtpOAuthApp.ps1`

Validates the full `SMTP.SendAsApp` flow end-to-end.

**What it tests:**

| Category | Tests |
|---|---|
| **Prerequisites** | Per-mailbox SMTP AUTH enabled; EXO service principal registered; `FullAccess` and `SendAs` permissions per mailbox |
| **Token** | Client credentials token acquisition; `aud` claim = `https://outlook.office365.com`; `SMTP.SendAsApp` in `roles` claim; token not expired |
| **SMTP (per mailbox)** | TCP connect to `smtp.office365.com:587`; STARTTLS negotiation; EHLO — XOAUTH2 advertised; `AUTH XOAUTH2` → 235 success; optional full message send → 250 End-of-DATA |

**Usage:**

```powershell
# Full test using config file
.\Test-SmtpOAuthApp.ps1 -ConfigFile ".\SmtpMailer-HR-smtp-config.json"

# Send to a specific recipient
.\Test-SmtpOAuthApp.ps1 `
    -ConfigFile    ".\SmtpMailer-HR-smtp-config.json" `
    -TestRecipient "admin@contoso.com"

# Auth validation only — do not send a message
.\Test-SmtpOAuthApp.ps1 `
    -ConfigFile   ".\SmtpMailer-HR-smtp-config.json" `
    -SkipSmtpSend
```

**Parameters:**

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-ConfigFile` | ❌* | — | Path to JSON config from `New-SmtpOAuthApp.ps1` (*recommended) |
| `-TenantId` | ❌* | — | Tenant ID (*required if no `-ConfigFile`) |
| `-AppId` | ❌* | — | Application client ID (*required if no `-ConfigFile`) |
| `-ClientSecret` | ❌* | — | Client secret (*required if no `-ConfigFile` and not using cert) |
| `-CertThumbprint` | ❌* | — | Certificate thumbprint (*required if no `-ConfigFile` and using cert) |
| `-Mailboxes` | ❌* | — | Mailbox addresses to test (*required if no `-ConfigFile`) |
| `-SkipSmtpSend` | ❌ | `$false` | Validate token and auth only — do not send a test message |
| `-TestRecipient` | ❌ | Sending mailbox | Recipient for test messages |
| `-TimeoutMs` | ❌ | `30000` | TCP connection timeout in milliseconds (5000–120000) |

**Exit codes:** `0` = all tests passed, `1` = one or more tests failed.

---

## Typical Workflow

### Graph API (Scoped Mailbox Access)

```
New-ScopedMailboxApp.ps1
        │
        └─ Wait 30–120 min for EXO permission propagation
                │
                └─ Test-ScopedMailboxApp.ps1
```

```powershell
# 1. Provision
.\New-ScopedMailboxApp.ps1 -AppName "MailApp-HR" -Mailboxes @("hr@contoso.com")

# 2. Wait for propagation, then test
.\Test-ScopedMailboxApp.ps1 `
    -AppName           "MailApp-HR" `
    -Mailboxes         @("hr@contoso.com") `
    -ConfigFile        ".\MailApp-HR-config.json" `
    -OutOfScopeMailbox "other@contoso.com"
```

### SMTP OAuth2

```
New-SmtpOAuthApp.ps1
        │
        └─ Wait 15–20 min for MAPI store propagation
                │
                └─ Test-SmtpOAuthApp.ps1
```

```powershell
# 1. Provision
.\New-SmtpOAuthApp.ps1 -AppName "SmtpMailer-HR" -Mailboxes @("hr@contoso.com")

# 2. Wait for propagation, then test
.\Test-SmtpOAuthApp.ps1 -ConfigFile ".\SmtpMailer-HR-smtp-config.json"
```

---

## Output Files

Both provisioning scripts write a JSON configuration file to the output directory.

| File | Contents |
|---|---|
| `<AppName>-config.json` | App ID, tenant ID, credential, custom attribute, scope name, MESG name — for Graph API flow |
| `<AppName>-smtp-config.json` | App ID, tenant ID, credential, EXO SP identity, SMTP host/port, per-mailbox permission results — for SMTP flow |

> **⚠️ Security:** Both files may contain a **client secret in plaintext**. Store them securely and do not commit them to source control. Add `*-config.json` to your `.gitignore`.

---

## Troubleshooting

### Graph API Flow

| Symptom | Cause | Fix |
|---|---|---|
| `Test-ServicePrincipalAuthorization` shows `InScope=True` for out-of-scope mailbox | Expected behaviour for non-exclusive scopes | Authoritative check is the Graph API 403/404 response |
| Graph API returns 200 for out-of-scope mailbox | Permissions not yet propagated | Wait 30–120 min and re-test |
| Role assignment `CustomResourceScope` is empty | Scope name not found at assignment time | Verify scope exists; remove assignment and re-run |

### SMTP OAuth2 Flow

| Symptom | Cause | Fix |
|---|---|---|
| Token `aud` is `https://graph.microsoft.com` | Wrong scope used | Use `https://outlook.office365.com/.default` |
| `SMTP.SendAsApp` missing from `roles` claim | Permission not added or consent not granted | Add `SMTP.SendAsApp` on the **EXO resource**, not Graph; grant admin consent |
| `535 5.7.3` at `AUTH XOAUTH2` | EXO service principal not registered, or wrong Object ID used | Re-run `New-ServicePrincipal` with the **Enterprise Application** Object ID |
| `XOAUTH2` not advertised after STARTTLS | Per-mailbox or org-level SMTP AUTH disabled | `Set-CASMailbox -SmtpClientAuthenticationDisabled $false`; check `Get-TransportConfig` |
| `430 STOREDRV MapiExceptionLogonFailed` | MAPI store not yet propagated | Wait 15–20 minutes and retry |

---

## References

- [Authenticate IMAP/POP/SMTP using OAuth — Microsoft Learn](https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth)
- [Application access policy for EWS and REST — Microsoft Learn](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access)
- [msgraph-sdk-powershell issue #3576](https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3576)
