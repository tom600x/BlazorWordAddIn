# Admin Setup Guide

One-time tenant-level setup that a **Global Administrator** or **Application Administrator** must perform before the add-in can be deployed. These steps are typically performed by the Entra/identity team and do not require access to Azure subscription resources.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Role | Entra **Application Administrator** or **Global Administrator** |
| Azure CLI 2.55+ | `winget install Microsoft.AzureCLI` — verify: `az --version` |
| PowerShell 7.x | Recommended; 5.1 also works |
| .NET 8 SDK | Required only if building/deploying the app yourself |
| Azure subscription | Needed only for infrastructure steps (separate from tenant admin steps) |

---

## 1. Create the Entra App Registration

### Script (automated)

```powershell
cd infra
.\app-registration.ps1 `
    -TenantId    "REPLACE_TENANT_ID" `
    -AppHostname "REPLACE_APP_NAME.azurewebsites.net"
```

The script performs all sub-steps below and prints the **Client ID** at completion. If your tenant already has an app registration with the same display name, it will reuse it rather than creating a duplicate.

### Manual

**1a. Sign in to the Azure portal**

Go to [https://portal.azure.com](https://portal.azure.com) and sign in with your Application Administrator or Global Administrator account.

**1b. Navigate to App Registrations**

In the left search bar, type **App registrations** and click it.  
Click **+ New registration**.

**1c. Fill in the registration form**

- **Name:** `BlazorWordAddIn` (or a name meaningful to your organization)
- **Supported account types:** `Accounts in this organizational directory only (Single tenant)`
- **Redirect URI:** Leave blank for now — you will add the SPA redirect URI separately in step 1f.

Click **Register**.

After registration, you land on the app's Overview page. Note the following values — you will need them in later steps:
- **Application (client) ID** — e.g. `4a8d9e5f-…`
- **Directory (tenant) ID** — e.g. `17784004-…`
- **Object ID** — (visible on the Overview page) — e.g. `7df82fe2-…`

**1d. Set the Application ID URI**

1. In the left menu, click **Expose an API**.
2. Next to **Application ID URI**, click **Set** (or **Edit** if one already exists).
3. Use the value `api://REPLACE_APP_NAME.azurewebsites.net/REPLACE_CLIENT_ID`.
   Example: `api://myapp.azurewebsites.net/4a8d9e5f-1234-5678-abcd-ef0123456789`
4. Click **Save**.

**1e. Add the `access_as_user` scope**

1. Still on **Expose an API**, click **+ Add a scope**.
2. Fill in:
   - **Scope name:** `access_as_user`
   - **Who can consent:** `Admins and users`
   - **Admin consent display name:** `Access Word Snippets on behalf of the user`
   - **Admin consent description:** `Allows the add-in to access the backend API on behalf of the signed-in user.`
   - **User consent display name:** `Access Word Snippets`
   - **User consent description:** `Allows the add-in to access the Word Snippets service on your behalf.`
   - **State:** `Enabled`
3. Click **Add scope**.
4. Note the full scope value: `api://HOSTNAME/CLIENT_ID/access_as_user` — you will need this in the manifest.

**1f. Pre-authorize Office client applications**

This step allows Office to silently obtain tokens without showing a consent prompt to each user.

1. Still on **Expose an API**, under **Authorized client applications**, click **+ Add a client application**.
2. Add each of the following Office client IDs **one at a time**. For each, paste the GUID, check the box next to `access_as_user`, and click **Add application**:

   | Application | Client ID |
   |---|---|
   | Office on the web | `57fb890c-0dab-4253-a5e0-7188c88b2bb4` |
   | Office (desktop) | `d3590ed6-52b3-4102-aeff-aad2292ab01c` |
   | Office Online Shell | `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` |
   | Office 365 Portal | `93d53678-613d-4013-afc3-0ebc4b444d30` |
   | Office on iPad/Android | `bc59ab01-8403-45c6-8796-ac3ef710b3e3` |
   | Microsoft Teams | `08e18876-6177-487e-b8b5-cf950c1e598c` |

   > **Tip:** Not all of these service principals may exist in your tenant. If a GUID is not recognized, skip it — the script uses only IDs that exist.

**1g. Add a SPA redirect URI (required for MSAL fallback)**

The add-in uses `auth.html` as the MSAL redirect page when Office SSO is unavailable (e.g. in Word Online). This URI must be registered as a **Single-page application** redirect, not a Web redirect.

1. In the left menu, click **Authentication**.
2. Scroll down to **Single-page application** (not "Web"). If this section is not visible, click **+ Add a platform**, choose **Single-page application**, enter the redirect URI shown below, and click **Configure**.
3. If the section already exists, click **Add URI**.
4. Enter: `https://REPLACE_APP_NAME.azurewebsites.net/auth.html`
5. Click **Save**.

> **Why `auth.html`?** In Word Online, `Office.auth.getAccessToken()` returns error `13006` because the Office client runs in a restricted iframe. The add-in falls back to opening `auth.html` via the Office Dialog API, running an MSAL login redirect, and posting the token back. The SPA redirect URI tells Entra where MSAL is allowed to redirect after login.

**1h. Add Microsoft Graph permissions**

1. In the left menu, click **API permissions**.
2. Click **+ Add a permission → Microsoft Graph → Delegated permissions**.
3. Search for and add each of these permissions:
   - `openid`
   - `profile`
   - `User.Read`
4. Click **Add permissions**.

Admin consent is granted separately in Step 2.

---

## 2. Grant Admin Consent

> This step requires **Global Administrator** or **Privileged Role Administrator** role.

### Portal

1. In the app registration, click **API permissions** in the left menu.
2. Verify you see `openid`, `profile`, and `User.Read` under **Microsoft Graph**.
3. Click **Grant admin consent for \<your tenant name\>**.
4. A confirmation dialog appears. Click **Yes**.
5. All permissions should show a green checkmark under **Status**.

### CLI

```bash
az ad app permission admin-consent --id REPLACE_ADDIN_CLIENT_ID
```

> Without admin consent, each user who opens the add-in receives a consent prompt the first time. For GCC deployments with many users, granting admin consent in advance is strongly recommended.

---

## 3. Add Optional Claims

Optional claims ensure that JWT access tokens issued by Entra contain the user identity fields the API needs (`preferred_username`, `upn`, `email`, `oid`). Without these, the API cannot determine the user's identity and returns a 401 or empty snippet list.

### Portal

1. In the app registration, click **Token configuration** in the left menu.
2. Click **+ Add optional claim**.
3. Select token type **Access**.
4. Check the boxes for: `email`, `preferred_username`, `upn`, `oid`
5. Click **Add**. If a dialog about enabling `profile` scope appears, check the box and confirm.
6. Repeat steps 2–5 for token type **ID token**.

### CLI (Graph API)

```powershell
$appObjectId = "REPLACE_APP_OBJECT_ID"   # Object ID (not Client ID) from the app registration Overview

$body = @{
    optionalClaims = @{
        accessToken = @(
            @{ name = "upn";                essential = $false }
            @{ name = "email";              essential = $false }
            @{ name = "preferred_username"; essential = $false }
            @{ name = "oid";               essential = $false }
        )
        idToken = @(
            @{ name = "upn";                essential = $false }
            @{ name = "email";              essential = $false }
            @{ name = "preferred_username"; essential = $false }
            @{ name = "oid";               essential = $false }
        )
    }
} | ConvertTo-Json -Depth 6

$body | Out-File -FilePath optional_claims.json -Encoding utf8

az rest `
    --method PATCH `
    --uri "https://graph.microsoft.com/v1.0/applications/$appObjectId" `
    --headers "Content-Type=application/json" `
    --body "@optional_claims.json"

Remove-Item optional_claims.json
```

---

## 4. Configure an Entra SQL Admin

> **Who:** Both the Entra admin and the Azure subscription admin (may be separate roles)

The SQL Server created by Bicep provisioning needs an Entra admin set before the post-deploy script can grant Managed Identity access. This cannot be set in the Bicep template without knowing the admin's object ID in advance.

### Portal

1. Go to [https://portal.azure.com](https://portal.azure.com).
2. Navigate to **SQL servers** → open `REPLACE_SQL_SERVER_NAME`.
3. In the left menu, under **Security**, click **Microsoft Entra ID**.
4. Click **Set admin**.
5. Search for and select the user or security group that will be the SQL admin.
6. Click **Select**, then **Save**.

### CLI

```bash
# Get the object ID of the user or group
az ad user show --id admin@contoso.com --query id --output tsv
# OR for a group:
az ad group show --group "SQL Admins" --query id --output tsv

# Set the Entra admin on the SQL Server
az sql server ad-admin create \
  --resource-group  REPLACE_RESOURCE_GROUP \
  --server-name     REPLACE_SQL_SERVER_NAME \
  --display-name    "REPLACE_DISPLAY_NAME" \
  --object-id       "REPLACE_OBJECT_ID"
```

---

## 5. Authentication Flow Reference

```
Word / Office client
  │
  ├─ Office.auth.getAccessToken()
  │    Succeeds on Word Desktop if Office is signed in with a work account.
  │    Returns error 13006 on Word Online (expected — see MSAL Fallback below).
  │    Token aud claim = CLIENT_ID GUID (NOT the App ID URI)
  │
  ├─ MSAL Fallback (Word Online / 13006)
  │    Add-in shows a "Sign In" button.
  │    User clicks → Office Dialog API opens auth.html.
  │    auth.html runs MSAL loginRedirect → user signs in → auth.html posts token back.
  │    Token aud claim = CLIENT_ID GUID
  │
  ▼
Azure AD (commercial, login.microsoftonline.com)
  Issues short-lived JWT (1 hour default)
  ▼
Blazor WASM task pane
  Attaches token as  Authorization: Bearer <jwt>
  ▼
ASP.NET Core API on App Service
  Microsoft.Identity.Web validates:
    • Signature (using Entra public keys)
    • Audience = CLIENT_ID GUID (configured in AzureAd:Audience)
    • Issuer = login.microsoftonline.com/<TenantId>/v2.0
    • Expiry
  Extracts UPN from preferred_username claim
  ▼
Azure SQL Database
  Managed Identity connection (no password)
```

> **GCC Note:** GCC (not GCC High) tenants use the **commercial** Azure AD endpoint `login.microsoftonline.com`. Do **not** configure `AzureAd:Instance` to `login.microsoftonline.us` unless your tenant is GCC High or DoD. See [GCC-Runbook.md](GCC-Runbook.md) for details.

---

## 6. Sharing Snippets Between Users

Snippets are private by default. To share a user's snippets with another user, insert a row into the `AllowedUsers` table:

```sql
-- Allow bob to see all of alice's snippets
INSERT INTO dbo.AllowedUsers (OwnerUpn, AllowedUpn)
VALUES (N'alice@contoso.com', N'bob@contoso.com');
```

The `AllowedUpn` user sees the owner's snippets in addition to their own. Sharing is one-directional; to share both ways, insert two rows.

To remove sharing:
```sql
DELETE FROM dbo.AllowedUsers
WHERE OwnerUpn = N'alice@contoso.com' AND AllowedUpn = N'bob@contoso.com';
```
