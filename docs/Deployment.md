# Deployment Guide

End-to-end steps to deploy the BlazorWordAddIn to Azure and publish via Centralized Deployment.

Two paths are provided for each step:
- **Script (automated)** — uses PowerShell scripts in `infra/`
- **Manual** — step-by-step portal/CLI instructions for teams where infra and app teams are separate

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Azure CLI 2.55+ | `winget install Microsoft.AzureCLI` — verify: `az --version` |
| Bicep CLI | `az bicep install` |
| .NET 8 SDK | https://dotnet.microsoft.com/download |
| Azure subscription | **Contributor** role on the target resource group |
| Entra role | **Application Administrator** or **Global Administrator** for app registration steps |
| sqlcmd | Ships with SQL Server tools; also available as `winget install Microsoft.SQLServerCommandLineUtils` |

Complete [Admin-Setup.md](Admin-Setup.md) before starting here.

---

## Step 1 — Create the Entra App Registration

> **Who:** Entra Application Administrator

### Script

```powershell
cd infra
.\app-registration.ps1 `
    -TenantId     "REPLACE_TENANT_ID" `
    -AppHostname  "REPLACE_APP_NAME.azurewebsites.net"
```

The script prints the **Client ID** at the end. Save it — you need it in Steps 2 and 6.

### Manual

See [Admin-Setup.md § 1](Admin-Setup.md#1-create-the-entra-app-registration) for the full manual walkthrough. The required outputs are:

- **Client ID** (a GUID, e.g. `4a8d9e5f-…`)
- **Application ID URI** (`api://HOSTNAME/CLIENT_ID`)
- **SPA redirect URI** configured as `https://HOSTNAME/auth.html`
- Optional claims enabled: `preferred_username`, `upn`, `email`, `oid`

---

## Step 2 — Populate the Parameters File

> **Who:** Infra / DevOps team

Edit `infra/parameters.bicepparam` and replace every `REPLACE_` value:

```bicep
using 'main.bicep'

param appName         = 'REPLACE_APP_NAME'         // App Service name (globally unique, 3–24 chars, lowercase)
param sqlServerName   = 'REPLACE_SQL_SERVER_NAME'  // SQL Server name (globally unique)
param sqlDatabaseName = 'SnippetsDb'
param addinClientId   = 'REPLACE_ADDIN_CLIENT_ID'  // Client ID GUID from Step 1
param tenantId        = 'REPLACE_TENANT_ID'        // Your Entra tenant ID
param location        = 'eastus'                   // Azure region (use 'usgovvirginia' for Azure Gov)
```

> **App name constraints:** 3–24 characters, lowercase letters and numbers only, globally unique across all of Azure. This becomes the default hostname: `REPLACE_APP_NAME.azurewebsites.net`.

---

## Step 3 — Deploy Azure Infrastructure

> **Who:** Infra / DevOps team (requires Azure subscription Contributor)

### Script

```powershell
cd infra
.\deploy.ps1 `
    -ResourceGroup "rg-word-snippets" `
    -Location      "eastus"
```

The script creates the resource group, deploys Bicep, publishes the app, and prints the hostname. Runtime: ~5–10 minutes.

### Manual

**3a. Create the resource group**

```bash
az login --tenant REPLACE_TENANT_ID
az account set --subscription REPLACE_SUBSCRIPTION_ID

az group create \
  --name     rg-word-snippets \
  --location eastus
```

**3b. Deploy the Bicep template**

```bash
cd infra

az deployment group create \
  --resource-group rg-word-snippets \
  --template-file  main.bicep \
  --parameters     parameters.bicepparam
```

Wait for the deployment to complete (usually 3–5 minutes). Note the `appHostname` output value; it is `REPLACE_APP_NAME.azurewebsites.net`.

**3c. Verify the App Service is running**

```bash
az webapp show \
  --name           REPLACE_APP_NAME \
  --resource-group rg-word-snippets \
  --query          "state" \
  --output         tsv
# Expected output: Running
```

**3d. Verify Managed Identity is enabled**

```bash
az webapp identity show \
  --name           REPLACE_APP_NAME \
  --resource-group rg-word-snippets \
  --query          "{principalId:principalId, tenantId:tenantId}" \
  --output         table
# principalId must not be empty
```

---

## Step 4 — Build and Publish the Application

> **Who:** Developer / DevOps team

### Script (included in deploy.ps1)

`deploy.ps1` already builds and publishes the application as part of Step 3.

### Manual

```powershell
# From the repo root
dotnet publish api\Api.csproj `
    --configuration Release `
    --output        publish_out `
    --self-contained false

# Zip the output
Compress-Archive -Path publish_out\* -DestinationPath app.zip -Force

# Deploy via zip deploy
az webapp deploy `
    --resource-group rg-word-snippets `
    --name           REPLACE_APP_NAME `
    --src-path       app.zip `
    --type           zip

# Confirm the deployment completed
az webapp deployment list-publishing-credentials `
    --name           REPLACE_APP_NAME \
    --resource-group rg-word-snippets \
    --query          "scmUri"
```

> After deployment, wait ~30 seconds and then browse to `https://REPLACE_APP_NAME.azurewebsites.net/api/me`. You should receive a `401 Unauthorized` — this is correct and confirms the app is running and JWT validation is active.

---

## Step 5 — Configure the SQL Entra Admin

> **Who:** Infra team (requires Entra privileges to set SQL Server admin)

> **Why:** The Bicep template provisions the SQL Server but cannot set the Entra admin automatically without knowing the admin's object ID in advance.

### Manual (Portal)

1. Go to the [Azure portal](https://portal.azure.com) → **SQL servers** → open `REPLACE_SQL_SERVER_NAME`.
2. In the left menu, click **Microsoft Entra ID** (under Security).
3. Click **Set admin**.
4. Search for and select the group or user account that will be the SQL admin (e.g. `SQL Admins` security group or your own account).
5. Click **Select**, then **Save**.

### Manual (CLI)

```bash
# Get the object ID of the intended SQL admin (user or group)
az ad user show --id admin@contoso.com --query id --output tsv
# OR for a group:
az ad group show --group "SQL Admins" --query id --output tsv

# Set the Entra admin on the SQL Server
az sql server ad-admin create \
  --resource-group  rg-word-snippets \
  --server-name     REPLACE_SQL_SERVER_NAME \
  --display-name    "REPLACE_DISPLAY_NAME" \
  --object-id       "REPLACE_OBJECT_ID"
```

---

## Step 6 — Grant Managed Identity Access to SQL

> **Who:** Database admin (must be the Entra SQL admin configured in Step 5)

### Script

```powershell
cd infra
.\post-deploy.ps1 `
    -ResourceGroup  "rg-word-snippets" `
    -AppName        "REPLACE_APP_NAME" `
    -SqlServerFqdn  "REPLACE_SQL_SERVER_NAME.database.windows.net" `
    -DatabaseName   "SnippetsDb"
```

### Manual

**6a. Acquire your Entra bearer token for Azure SQL**

```powershell
$tokenJson   = az account get-access-token --resource https://database.windows.net/ | ConvertFrom-Json
$accessToken = $tokenJson.accessToken
```

**6b. Connect to the database and create the contained user**

```powershell
# Install the SqlServer module if you don't have it
Install-Module -Name SqlServer -Force -AllowClobber -Scope CurrentUser

Invoke-Sqlcmd `
    -ServerInstance "REPLACE_SQL_SERVER_NAME.database.windows.net" `
    -Database       "SnippetsDb" `
    -AccessToken    $accessToken `
    -Query @"
-- Create a contained database user mapped to the App Service Managed Identity.
-- The name must match the App Service name exactly (it is the MI display name).
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = 'REPLACE_APP_NAME')
BEGIN
    CREATE USER [REPLACE_APP_NAME] FROM EXTERNAL PROVIDER;
END

-- Grant read/write access
ALTER ROLE db_datareader ADD MEMBER [REPLACE_APP_NAME];
ALTER ROLE db_datawriter ADD MEMBER [REPLACE_APP_NAME];

PRINT 'Done.';
"@
```

**6c. Verify**

```powershell
Invoke-Sqlcmd `
    -ServerInstance "REPLACE_SQL_SERVER_NAME.database.windows.net" `
    -Database       "SnippetsDb" `
    -AccessToken    $accessToken `
    -Query "SELECT name, type_desc FROM sys.database_principals WHERE type_desc IN ('EXTERNAL_USER','EXTERNAL_GROUP');"
# REPLACE_APP_NAME should appear in the results
```

---

## Step 7 — Create the Database Schema

> **Who:** Database admin

```powershell
$tokenJson   = az account get-access-token --resource https://database.windows.net/ | ConvertFrom-Json
$accessToken = $tokenJson.accessToken

# Apply the schema (creates tables if they do not exist; safe to run multiple times)
sqlcmd `
    -S "REPLACE_SQL_SERVER_NAME.database.windows.net" `
    -d SnippetsDb `
    -G `
    -P $accessToken `
    -i sql\schema.sql

# Optional: load sample snippets
sqlcmd `
    -S "REPLACE_SQL_SERVER_NAME.database.windows.net" `
    -d SnippetsDb `
    -G `
    -P $accessToken `
    -i sql\seed.sql
```

> **Note:** The `-G` flag tells sqlcmd to use Azure Active Directory authentication. `-P $accessToken` passes the bearer token obtained above. If your version of sqlcmd does not support `-G -P <token>`, use the SSMS "Connect with Azure Active Directory" option instead.

---

## Step 8 — Update the Add-in Manifest

> **Who:** Developer

Open `add-in/manifest.xml` and replace all `REPLACE_` tokens:

| Token | Replace with | Example |
|---|---|---|
| `REPLACE_ADDIN_MANIFEST_GUID` | Any new random GUID | `7f3a1c2d-4e5f-6789-abcd-ef0123456789` |
| `REPLACE_ADDIN_CLIENT_ID` | Entra app **Client ID** from Step 1 | `4a8d9e5f-…` |
| `REPLACE_APP_HOSTNAME` | App Service default hostname | `myapp.azurewebsites.net` |

To generate a GUID in PowerShell:
```powershell
[System.Guid]::NewGuid().ToString()
```

**Validate the manifest** before deploying to avoid silent failures:
- Upload it via Word Desktop: **Insert → Add-ins → My Add-ins → Upload My Add-in**
- Or use the [Office Add-in Validator](https://github.com/OfficeDev/office-addin-manifest): `npx office-addin-manifest validate add-in/manifest.xml`

---

## Step 9 — Centralized Deployment (M365 Admin)

> **Who:** M365 Global Admin or Exchange Admin  
> GCC organizations **must** use Centralized Deployment — the Office Store is not available in GCC.

**Portal steps (most common):**

1. Open a browser and sign in to [https://admin.microsoft.com](https://admin.microsoft.com) with your Global Admin or Exchange Admin account.
2. In the left navigation, go to **Settings → Integrated Apps**.
3. Click **Deploy Add-in** (top of the page) or **+ Add apps**.
4. Select **Upload custom app**, then choose **I have a manifest file (.xml) on my device**.
5. Click **Browse**, navigate to `add-in/manifest.xml` in this repository, and click **Upload**.
6. Review the permissions listed. Click **Next**.
7. Under **Who gets access**, choose one of:
   - **Everyone** — all users in the tenant receive the add-in
   - **Specific users/groups** — recommended for pilot rollouts; start here
   - **Just me** — useful for personal testing only
8. If you chose **Specific users/groups**, click **Add users**, search for and select the target groups or users, then click **Add**.
9. Review the deployment summary and click **Deploy**.
10. Click **Done** to close the wizard.

> **Propagation time:** 15–30 minutes for initial rollout. Changes to an existing deployment can take up to 2 hours. Users must **close and reopen all Office applications** for the add-in to appear.

**Verify deployment status:**

1. In the M365 Admin Center, go to **Settings → Integrated Apps**.
2. Find the **BlazorWordAddIn** entry. The status should be **Deployed**.
3. If status shows **Failed**, click the entry to see error details. Common issues:
   - Manifest validation error — re-check the XML for malformed GUIDs or URLs
   - Permission consent not granted — see [Admin-Setup.md § Admin Consent](Admin-Setup.md#2-grant-admin-consent)

---

## Step 10 — Smoke Test

> **Who:** Anyone in the deployment scope

**Word Desktop:**
1. Close all Office applications completely (File → Exit, not just close the window).
2. Wait 2 minutes, then reopen Microsoft Word.
3. Open any `.docx` document (the task pane cannot operate on a blank session).
4. Look for a **Word Snippets** button in the **Home** tab of the ribbon. If not visible after 30 minutes, check the Centralized Deployment status.
5. Click **Word Snippets** — the task pane opens.
6. The pane should display your name/UPN at the top, confirming SSO authentication succeeded.
7. If a **Sign In** button is shown instead, click it — a browser dialog window will open asking for Microsoft credentials. This is the MSAL fallback and is expected in Word Online (see Note below).
8. After sign-in, the snippet list should load. Click **Insert** on any text snippet and confirm it appears in the document.

**Word on the Web (Word Online):**
1. Open [https://office.com](https://office.com) and open or create a Word document.
2. From the task pane (Insert → Add-ins → Admin Managed), open **Word Snippets**.
3. A **Sign In** button will be shown — click it. A popup window opens for MSAL authentication.
4. Sign in with your Microsoft 365 work account. The window closes automatically.
5. The snippet list loads.

> **Note on Sign In button:** In Word Online, Office SSO (`getAccessToken`) returns error `13006` due to Word Online's iframe security model. This is **expected and handled** — the add-in automatically falls back to showing a Sign In button that triggers MSAL authentication via a popup dialog. The experience is seamless after the first sign-in; the token is cached for the browser session.

---

## Environment Overview

| Resource | Azure Service | SKU | Notes |
|---|---|---|---|
| Add-in UI + API | App Service | B1 Linux (.NET 8) | Single App Service hosts both UI and API |
| Database | Azure SQL Database | S0 (10 DTUs) | Entra auth only; no SQL passwords |
| Auth | Managed Identity (system-assigned) | — | Credentials managed automatically by Azure |
| Identity provider | Azure AD (commercial) | — | GCC uses commercial endpoint (`login.microsoftonline.com`) |

---

## App Service Configuration Values (Reference)

These values are set by Bicep automatically but are useful to verify in the portal (**App Service → Configuration → Application settings**):

| Setting | Expected value |
|---|---|
| `ASPNETCORE_ENVIRONMENT` | `Production` |
| `AzureAd__Instance` | `https://login.microsoftonline.com/` |
| `AzureAd__TenantId` | Your tenant GUID |
| `AzureAd__ClientId` | Entra app Client ID GUID |
| `AzureAd__Audience` | **The Client ID GUID** (same as `ClientId`) — NOT the App ID URI |
| `ConnectionStrings__DefaultConnection` | `Server=tcp:SERVERNAME.database.windows.net,...;Authentication=Active Directory Default;...` |

> **Why does `Audience` equal `ClientId`?** MSAL's SPA token flow (`PublicClientApplication.loginRedirect`) places the application's client GUID in the `aud` claim of the access token — it does **not** use the full App ID URI. The backend must match what the token actually contains. This is a known difference from the Office SSO flow and is documented in the code.

---

## Rollback / Teardown

```bash
# Delete all Azure resources (irreversible — data will be lost)
az group delete --name rg-word-snippets --yes --no-wait

# Remove app registration
az ad app delete --id REPLACE_ADDIN_CLIENT_ID
```

To remove the add-in from users without deleting the backend, go to **M365 Admin Center → Integrated Apps**, find the add-in, and click **Remove**.
