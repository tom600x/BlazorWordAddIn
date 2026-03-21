# BlazorWordAddIn

A Microsoft Word Task Pane Office Add-in with a Blazor WebAssembly UI, targeting US Government Community Cloud (GCC) environments.

## Architecture

```
Word Desktop / Word Online
        │
        ├─ Office.auth.getAccessToken()  (Word Desktop — succeeds silently)
        └─ Error 13006 → Sign In button  (Word Online — MSAL dialog fallback)
                │
                ▼ JWT  (aud = CLIENT_ID GUID)
┌───────────────────────────────┐
│  Task Pane                    │
│  Blazor WASM UI               │
│  - Text Snippets tab          │
│  - Image Snippets tab         │
│  - auth.html (MSAL fallback)  │
└──────────────┬────────────────┘
               │ Authorization: Bearer <jwt>
               ▼
┌───────────────────────────────┐
│  ASP.NET Core API             │
│  (Azure App Service)          │
│  - Validates JWT              │
│     aud = CLIENT_ID GUID      │
│     iss = AAD v2 endpoint     │
│  - Extracts UPN from claims   │
└──────────────┬────────────────┘
               │ Managed Identity (no password)
               ▼
┌───────────────────────────────┐
│  Azure SQL Database           │
│  - TextSnippets               │
│  - ImageSnippets              │
│  - AllowedUsers               │
└───────────────────────────────┘
```

## Repository Layout

```
BlazorWordAddIn/
├── add-in/          Office add-in manifest + auth.html (MSAL fallback page)
├── blazor-ui/       Blazor WebAssembly project (task pane UI)
├── api/             ASP.NET Core Web API (data + auth) — also hosts Blazor static files
├── infra/           Bicep templates + PowerShell deploy/registration scripts
├── sql/             Database schema + seed data
└── docs/            Admin, developer, deployment, and GCC documentation
```

## Quick Start (Local Development)

### Prerequisites

- .NET 8 SDK
- Visual Studio 2022 17.8+ or VS Code with C# Dev Kit
- Azure subscription (for SQL and App Service)
- Microsoft 365 tenant (GCC or commercial) with Application Administrator access for app registration

### Step 1 — Clone

```bash
git clone <repo-url>
cd BlazorWordAddIn
```

### Step 2 — Register the Entra app

Follow [docs/Admin-Setup.md](docs/Admin-Setup.md) or run the automation script:

```powershell
cd infra
.\app-registration.ps1 -TenantId "REPLACE_TENANT_ID" -AppHostname "REPLACE_APP_NAME.azurewebsites.net"
```

The script creates the app registration, sets the App ID URI, adds the `access_as_user` scope, pre-authorizes Office client IDs, registers the SPA redirect URI for `auth.html`, and adds optional claims.

### Step 3 — Configure local secrets

```powershell
cd api
dotnet user-secrets set "AzureAd:TenantId"  "REPLACE_TENANT_ID"
dotnet user-secrets set "AzureAd:ClientId"  "REPLACE_ADDIN_CLIENT_GUID"
dotnet user-secrets set "AzureAd:Audience"  "REPLACE_ADDIN_CLIENT_GUID"
```

> **Audience must be the Client ID GUID.** Do not use the full App ID URI (`api://HOSTNAME/GUID`). MSAL tokens set `aud = CLIENT_GUID`, and the API validates against this value. Using the App ID URI causes `401 IDX10214` errors on every request.

### Step 4 — Apply the database schema

```powershell
# Using sqlcmd with SQL Server LocalDB:
sqlcmd -S "(localdb)\mssqllocaldb" -d master   -i sql\schema.sql
sqlcmd -S "(localdb)\mssqllocaldb" -d BlazorWordAddIn -i sql\seed.sql
```

### Step 5 — Trust the dev certificate and run

```powershell
dotnet dev-certs https --trust

# Start the API (also serves Blazor WASM):
cd api
dotnet run
```

### Step 6 — Sideload the add-in

1. Update `add-in/manifest.xml` — change `SourceLocation` to `https://localhost:7000`.
2. Open Word → **Insert → Add-ins → My Add-ins → Upload My Add-in** → browse to `manifest.xml`.
3. Click the **Word Snippets** button in the Home ribbon.

See [docs/Developer-Setup.md](docs/Developer-Setup.md) for full local dev instructions, including MSAL fallback testing in Word Online.

---

## Cloud Deployment

See [docs/Deployment.md](docs/Deployment.md) for full step-by-step instructions with both script and manual portal procedures.

```powershell
cd infra
.\deploy.ps1 `
    -ResourceGroup  "rg-snippets" `
    -Location       "eastus" `
    -AppName        "snippets-app" `
    -SqlServerName  "snippets-sql"
```

---

## GCC Deployment

See [docs/GCC-Runbook.md](docs/GCC-Runbook.md) for:
- GCC vs GCC High vs DoD endpoint differences
- Centralized Deployment walkthrough (M365 Admin Center)
- MSAL fallback behavior in GCC Word Online
- Azure Government (Option B) deployment notes

---

## Authentication Notes

| Scenario | Token source | `aud` claim | Sign-in experience |
|---|---|---|---|
| Word Desktop (M365 signed in) | `Office.auth.getAccessToken()` | Client ID GUID | Silent — no prompt |
| Word Online | Error 13006 → MSAL dialog | Client ID GUID | User clicks **Sign In** button |
| Local dev (any browser) | MSAL dialog via auth.html | Client ID GUID | User clicks **Sign In** button |

The `Audience` setting in `appsettings.json` / App Service app settings must equal the **Client ID GUID**.  
The full App ID URI (`api://HOSTNAME/GUID`) appears only in the manifest's `<Resource>` element, not in `Audience`.

---

## Replace These Values

Before running, replace the following placeholders across the codebase:

| Placeholder | Description | Where |
|---|---|---|
| `REPLACE_ADDIN_CLIENT_GUID` | Entra app registration **Client ID** (GUID) | `manifest.xml`, `appsettings.json`, `parameters.bicepparam`, user-secrets |
| `REPLACE_TENANT_ID` | Azure / M365 **Tenant ID** (GUID) | `appsettings.json`, `parameters.bicepparam`, scripts |
| `REPLACE_APP_NAME` | Azure App Service name (e.g. `snippets-app`) | `manifest.xml`, `appsettings.json`, `parameters.bicepparam` |
| `REPLACE_APP_HOSTNAME` | App Service hostname (e.g. `snippets-app.azurewebsites.net`) | `manifest.xml`, `app-registration.ps1` |
| `REPLACE_SQL_SERVER_NAME` | Azure SQL Server name | `appsettings.json`, `parameters.bicepparam` |
| `REPLACE_DB_NAME` | SQL Database name | `appsettings.json`, `parameters.bicepparam` |
| `REPLACE_RESOURCE_GROUP` | Azure resource group name | deploy scripts |

---

## Security Notes

- No secrets are stored in source code. Sensitive values use App Service configuration, Azure Key Vault, or `dotnet user-secrets` for local dev.
- The add-in authenticates via Office SSO or MSAL dialog — the JWT is validated by the API on every request; the browser never accesses the database.
- Managed Identity is used for Azure SQL access — no SQL passwords in config, no rotation required.
- All traffic is TLS 1.2+; HTTP is redirected to HTTPS by the App Service.
- GCC customers: all authentication endpoints use commercial Azure AD (`login.microsoftonline.com`). No special network configuration required beyond what GCC already allows.

---

## Solution Structure

```
BlazorWordAddIn.sln
├── blazor-ui/BlazorUI.csproj      (Microsoft.NET.Sdk.BlazorWebAssembly)
└── api/Api.csproj                 (Microsoft.NET.Sdk.Web — hosts Blazor WASM + API endpoints)
```

The API project hosts the Blazor WASM static files, so a single Azure App Service serves both the task pane UI and the API.

---

## Documentation Index

| Document | Audience | Contents |
|---|---|---|
| [docs/Admin-Setup.md](docs/Admin-Setup.md) | Entra Application Administrator | App registration, optional claims, SPA redirect URI, SQL Entra admin, admin consent |
| [docs/Deployment.md](docs/Deployment.md) | Infrastructure / DevOps team | Bicep deployment, App Service config, post-deploy SQL, M365 Centralized Deployment |
| [docs/Developer-Setup.md](docs/Developer-Setup.md) | Developer | Local dev setup, user secrets, sideloading, MSAL fallback testing |
| [docs/Troubleshooting.md](docs/Troubleshooting.md) | Any | Error code reference, 401 diagnostics, audience mismatch fix, 500 errors |
| [docs/GCC-Runbook.md](docs/GCC-Runbook.md) | GCC M365 Admin / helpdesk | GCC-specific deployment, centralized deployment steps, network requirements |
