# Developer Setup Guide

This document explains how to clone, configure, run, and sideload the add-in locally for iterative development.

---

## Prerequisites

| Tool | Version | Install |
|---|---|---|
| .NET SDK | 8.x | https://dotnet.microsoft.com/download |
| Visual Studio 2022 | 17.8+ **or** VS Code | https://visualstudio.microsoft.com |
| Azure CLI | 2.55+ | `winget install Microsoft.AzureCLI` |
| Office Desktop | Microsoft 365 E3/G3+ build 16.0.14326+ | Required for Office SSO |
| SQL Server Express or LocalDB | Any | Ships with VS 2022 |

---

## 1. Clone and Open

```bash
git clone https://github.com/YOUR_ORG/BlazorWordAddIn.git
cd BlazorWordAddIn

# Open in VS Code
code .

# OR open in Visual Studio
start BlazorWordAddIn.sln
```

---

## 2. Configure User Secrets (API)

Do **not** edit `appsettings.Development.json` directly for sensitive values. Use .NET User Secrets instead:

```bash
cd api

dotnet user-secrets set "AzureAd:TenantId"  "REPLACE_TENANT_ID"
dotnet user-secrets set "AzureAd:ClientId"  "REPLACE_ADDIN_CLIENT_GUID"
dotnet user-secrets set "AzureAd:Audience"  "REPLACE_ADDIN_CLIENT_GUID"
```

> **Important — Audience value:** Set `Audience` to the **Client ID GUID** only (e.g. `4a8d9e5f-1234-5678-abcd-ef0123456789`).  
> Do **not** use the full App ID URI (`api://localhost:7001/CLIENT_ID`) — this will cause `401 IDX10214` errors.  
> MSAL issues tokens where `aud = CLIENT_GUID`, and Microsoft.Identity.Web validates against the configured `Audience`.

The `appsettings.Development.json` file already provides non-sensitive defaults:
- LocalDB connection string for `DefaultConnection`
- CORS allowed origins for `https://localhost:7000`

---

## 3. Create the Local Database

Run the schema script against LocalDB before starting the API:

```powershell
# Option A: sqlcmd (ships with SQL Server tools)
sqlcmd -S "(localdb)\mssqllocaldb" -d master -i sql\schema.sql

# Load seed data (optional — adds sample snippets for your dev UPN):
sqlcmd -S "(localdb)\mssqllocaldb" -d BlazorWordAddIn -i sql\seed.sql
```

> The seed data uses a placeholder UPN. Edit `sql\seed.sql` to replace it with your own sign-in UPN so the snippets show up when you test.

---

## 4. Trust the Developer Certificate

```powershell
dotnet dev-certs https --trust
```

Office add-ins require HTTPS with a valid (trusted) certificate. Run this once after a fresh clone; it is not needed again unless you switch machines.

---

## 5. Start the API

```bash
cd api
dotnet run
```

The API starts on `https://localhost:7001`.

Verify: open `https://localhost:7001/api/me` in a browser — you should get `401 Unauthorized`. This is correct — auth is required.

---

## 6. Update the Blazor App Settings (Dev)

Open `blazor-ui/wwwroot/appsettings.Development.json` and confirm:

```json
{
  "ApiBaseUrl": "https://localhost:7001"
}
```

---

## 7. Run the Blazor Dev Server

```bash
cd blazor-ui
dotnet run
# Starts on https://localhost:7000
```

---

## 8. Sideload the Add-in in Word Desktop

Before sideloading, update `add-in/manifest.xml` to use the local dev URL:

1. Open `add-in/manifest.xml`.
2. Find every `<SourceLocation DefaultValue="https://tomwordaddin.azurewebsites.net/...">` attribute.
3. Replace the hostname with `https://localhost:7000` for local dev.

Then sideload:

1. Open **Microsoft Word**.
2. **Insert** → **Add-ins** → **My Add-ins** → **Upload My Add-in**.
3. Browse to `add-in/manifest.xml` and click **Upload**.
4. The **Word Snippets** button appears in the **Home** ribbon.

---

## 9. Authentication in Local Dev

### Office SSO (getAccessToken)

When you open the add-in task pane in Word Desktop:

- If Word is signed in with your organizational account, `Office.auth.getAccessToken()` succeeds silently and returns a JWT.
- Decode this JWT at https://jwt.ms and check:
  - `aud` = your Client ID GUID (NOT `api://...`)
  - `preferred_username` = your UPN
  - `tid` = your Tenant ID

### MSAL Fallback (Word Online / SSO Error 13006)

In Word Online, `getAccessToken()` returns error `13006`. This is **expected behavior** — Office SSO does not function in Word Online due to iframe sandbox restrictions.

The add-in handles this automatically:

1. Detects error 13006.
2. Displays a **Sign In** button in the task pane.
3. When the user clicks **Sign In**, the Office Dialog API opens `https://localhost:7000/auth.html`.
4. `auth.html` runs MSAL's `loginRedirect`. The user signs in in the dialog.
5. `auth.html` posts the token back to the task pane.

**Prerequisite for local MSAL fallback:** Add `https://localhost:7000/auth.html` as a **Single-page application** redirect URI in the Entra app registration. See [Admin-Setup.md](Admin-Setup.md) step 1g.

---

## 10. Sideload in Word on the Web (Word Online)

1. Open any Word document in Office 365 (web).
2. **Insert** → **Add-ins** → **My Add-ins** → **Upload My Add-in** → **Browse** → select `add-in/manifest.xml`.
3. Accept any prompts and reload the page.
4. Click the **Sign In** button when prompted — the MSAL fallback dialog opens.

---

## 11. Running Both Projects at Once

### VS Code

Add a `.vscode/launch.json` compound launch that starts both the API and Blazor-UI simultaneously:

```json
{
  "version": "0.2.0",
  "compounds": [
    {
      "name": "API + Blazor",
      "configurations": ["Launch API", "Launch Blazor"]
    }
  ],
  "configurations": [
    {
      "name": "Launch API",
      "type": "coreclr",
      "request": "launch",
      "preLaunchTask": "build-api",
      "program": "${workspaceFolder}/api/bin/Debug/net8.0/Api.dll",
      "cwd": "${workspaceFolder}/api"
    },
    {
      "name": "Launch Blazor",
      "type": "coreclr",
      "request": "launch",
      "preLaunchTask": "build-blazor",
      "program": "${workspaceFolder}/blazor-ui/bin/Debug/net8.0/BlazorUI.dll",
      "cwd": "${workspaceFolder}/blazor-ui"
    }
  ]
}
```

### Visual Studio

- Right-click solution → **Set Startup Projects → Multiple startup projects**.
- Set both `BlazorUI` and `Api` to **Start**.

---

## 12. Useful Commands

| Task | Command |
|---|---|
| Run API | `cd api && dotnet run` |
| Run Blazor WASM | `cd blazor-ui && dotnet run` |
| Build entire solution | `dotnet build BlazorWordAddIn.sln` |
| List user secrets | `cd api && dotnet user-secrets list` |
| Add EF migration | `cd api && dotnet ef migrations add <Name>` |
| Apply migrations | `cd api && dotnet ef database update` |

---

## Troubleshooting (Quick Reference)

| Issue | Fix |
|---|---|
| `IDX10214: Audience validation failed` | `Audience` in user secrets must be the GUID, not `api://...`. See Step 2. |
| SSO error 13006 in Word Online | Expected — click the **Sign In** button. See Step 9. |
| SSO error 13001 | Sign in to Office with your work/school account (File → Account → Sign In) |
| SSO error 13002 | Run `app-registration.ps1`; grant admin consent in Entra portal |
| CORS error in browser | Verify `Cors:AllowedOrigins` in appsettings includes your dev URL |
| Add-in not loading | Check manifest `SourceLocation` matches the running Blazor URL; check HTTPS cert |
| Snippets list is empty | Check seed.sql UPN matches your sign-in UPN; check SQL schema was applied |

See [Troubleshooting.md](Troubleshooting.md) for full details.
