# Troubleshooting Guide

Reference for diagnosing and resolving common issues with the BlazorWordAddIn.

---

## Office SSO (`getAccessToken`) Errors

The following error codes are returned by `Office.auth.getAccessToken()` and surfaced in the task pane.

| Code | Name | Cause | Resolution |
|---|---|---|---|
| **13001** | User not signed in | The Office host is not signed in with an Entra work/school account | Tell the user: File → Account → Sign in with your work account |
| **13002** | User has not consented | User (or admin) has not consented to the required permissions | Run `infra/app-registration.ps1` then grant admin consent in Entra portal |
| **13003** | Unsupported user type | Microsoft Account (personal) is not supported; only Entra work/school | User must use their organizational account |
| **13005** | Invalid grant | The add-in's app registration is misconfigured | Verify `WebApplicationInfo` in manifest; re-run `app-registration.ps1` |
| **13006** | Client error (Word Online) | Office SSO is not available in Word Online's restricted iframe | **Expected behavior** — see MSAL Fallback section below |
| **13007** | Silent token failed | Add-in requires user interaction to complete auth | Retry with `allowSignInPrompt: true` in `getAccessToken` options |
| **13008** | SSO retrieval in progress | Another SSO operation is currently running | Wait and retry; do not call `getAccessToken` before the previous call resolves |
| **13010** | Browser issue | Webview issue inside the Office host | Try the add-in in a different Office application or on Word on the web |
| **13012** | Invalid app ID or manifest | The add-in is not registered correctly or manifest GUID mismatch | Verify `<Id>` in manifest matches the Entra app client ID; re-sideload the manifest |

---

## Error 13006 in Word Online — MSAL Fallback

**Symptom:** The task pane displays a **Sign In** button rather than the snippets list. The browser console may show `getAccessToken error 13006`.

**This is expected behavior.** In Word Online, the Office client runs inside an iframe with sandbox restrictions that prevent silent SSO. Error 13006 is not a misconfiguration.

**How the fallback works:**

1. The add-in detects error 13006 from `getAccessToken`.
2. A **Sign In** button is displayed.
3. The user clicks **Sign In**.
4. The Office Dialog API opens `auth.html` in a separate popup window.
5. `auth.html` runs MSAL's `loginRedirect`. The user completes sign-in in the popup.
6. `auth.html` posts the JWT token back to the task pane.
7. The task pane receives the token and loads the snippets list.

**If the Sign In button click does nothing or the dialog stays blank:**
- Confirm `https://YOUR_APP_NAME.azurewebsites.net/auth.html` is registered as a **Single-page application** redirect URI in the Entra app registration (not a Web redirect). See [Admin-Setup.md](Admin-Setup.md) step 1g.
- Confirm pop-ups are not blocked by the browser or a corporate proxy.

---

## HTTP 401 Unauthorized

**Symptom:** The task pane shows "Error loading snippets" or "Unauthorized" immediately after obtaining a token.

### Most common cause: Audience (aud) mismatch

This is the #1 cause of 401 errors. To diagnose:

1. Open the browser DevTools console (F12).
2. Look for the error message from the `X-Auth-Error` response header. An audience mismatch looks like:
   ```
   IDX10214: Audience validation failed.
   Audiences: 'c7806dca-7fab-4e78-bf17-d3f977b3eea4'.
   Did not match: validationParameters.ValidAudience: 'api://myapp.azurewebsites.net/c7806dca-...'
   ```
3. Decode the JWT at https://jwt.ms and inspect the `aud` claim.

**Key insight:** MSAL (used by both Office SSO and the `auth.html` fallback) issues tokens where `aud = CLIENT_ID GUID`. The `Audience` in the API's `appsettings.json` must be set to **just the GUID**, not the full `api://HOSTNAME/GUID` URI.

**Correct configuration:**
```json
"AzureAd": {
  "Audience": "c7806dca-7fab-4e78-bf17-d3f977b3eea4"
}
```

**Incorrect (causes IDX10214):**
```json
"AzureAd": {
  "Audience": "api://tomwordaddin.azurewebsites.net/c7806dca-7fab-4e78-bf17-d3f977b3eea4"
}
```

For Bicep deployments: verify `infra/modules/appservice.bicep` sets `AzureAd__Audience` to `addinClientId` (the GUID parameter), not `api://${appName}.azurewebsites.net/${addinClientId}`.

### Other 401 checklist

| Check | How to verify |
|---|---|
| `ClientId` in appsettings matches Entra app | Compare `appsettings.json:AzureAd:ClientId` with the Client ID in the Azure portal app registration |
| Token is not expired | Check `exp` claim in jwt.ms — must be in the future |
| `TenantId` in appsettings matches token | Compare `appsettings.json:AzureAd:TenantId` with `tid` claim in jwt.ms |
| Manifest `<Resource>` URI is correct | Must be `api://HOSTNAME/CLIENT_GUID` (used by Office SSO scoping, not by `Audience` validation) |

---

## HTTP 403 Forbidden

**Symptom:** Task pane loads but data endpoints return 403.

**Likely causes:**
- The JWT was issued for a different tenant (`tid` mismatch).
- The user's account is from a guest/external tenant.
- `TenantId` in `appsettings.json` is set to a specific GUID and the token's tenant does not match.

**Fix:** Check `api/appsettings.json` → `AzureAd:TenantId`. If you want to support multiple tenants, set it to `common` (not recommended for GCC — pin to your tenant ID).

---

## HTTP 500 Internal Server Error

**Symptom:** API returns HTTP 500. The task pane shows a generic error with no detail.

**Diagnosing the exception:**

The API is configured to return the exception message in the response body when a 500 occurs (for diagnostic purposes). Check the response body in browser DevTools (Network tab → the failed request → Response) for the full exception text.

Additionally, check App Service logs:

```bash
az webapp log tail \
  --name           REPLACE_APP_NAME \
  --resource-group REPLACE_RESOURCE_GROUP
```

### Common 500 causes

| Cause | Diagnostic sign | Fix |
|---|---|---|
| SQL connection failure | `SqlException: Login failed` or `Cannot open server` | See SQL connection failures section below |
| Managed Identity not configured | `ManagedIdentityCredential: MSI is not enabled` | Enable system-assigned identity: `az webapp identity assign` |
| Missing optional claims | 500 on `/api/me` | Ensure `preferred_username` or `upn` optional claims are added in Entra (see [Admin-Setup.md](Admin-Setup.md) step 3) |
| App Service plan stopped | Cold start timeout | Check plan state; check memory/CPU quotas in Azure portal |

---

## CORS Errors

**Symptom:** Browser console shows `Access-Control-Allow-Origin` error when the task pane calls the API.

**Fix:**
1. Confirm `Cors:AllowedOrigins` in `appsettings.json` contains the exact origin of the task pane (no trailing slash).
2. If the add-in and API are on the **same App Service** (default setup), CORS is not needed — the origin matches.
3. For dev: add `https://localhost:7000` (Blazor dev server port) to `appsettings.Development.json`.

> The API uses `.WithExposedHeaders("X-Auth-Error")` in its CORS policy to allow the task pane to read the `X-Auth-Error` diagnostic header on 401 responses.

---

## Azure SQL Connection Failures

**Symptom:** API starts but any database operation throws `SqlException` or task pane shows a 500 error.

**Step-by-step:**

1. **Check that post-deploy.ps1 was run:**
   ```sql
   -- Connect to the database as the Entra SQL admin and run:
   SELECT name, type_desc FROM sys.database_principals
   WHERE type_desc IN ('EXTERNAL_USER', 'EXTERNAL_GROUP');
   ```
   The App Service name should appear.

2. **Verify the connection string in App Settings** (Azure portal → App Service → Configuration):
   - Must contain `Authentication=Active Directory Default`
   - Server name must end in `.database.windows.net`
   - Database name must match what was created

3. **Check SQL firewall** (Azure portal → SQL Server → Networking):
   - `Allow Azure services and resources to access this server` must be **ON**

4. **Check Managed Identity is enabled:**
   ```bash
   az webapp identity show \
     --name           REPLACE_APP_NAME \
     --resource-group REPLACE_RESOURCE_GROUP
   ```
   `principalId` must be non-null.

---

## Add-in Not Appearing in Word Ribbon

**Symptom:** After Centralized Deployment, the **Word Snippets** button does not appear in the Home tab.

**Checklist:**
1. Check M365 Admin Center → **Integrated Apps** — confirm status is **Deployed** (not Failed or Pending).
2. User must be in the targeted deployment scope (group membership). Verify via `az ad group member check --group GRP_NAME --member-id USER_OID`.
3. User must **close all Office apps and relaunch** — propagation requires a fresh session.
4. Propagation time can be up to **30 minutes** for new deployments, longer for tenant-wide.
5. For Word Desktop, check the Office build meets the minimum requirement (build ≥ 16.0.14326).

---

## Word Insertion Fails (Error on Insert Button)

**Symptom:** User clicks **Insert** on a snippet and sees "Insertion failed" or nothing happens.

**Checklist:**
1. A Word document must be **open and active** — the task pane cannot insert into an empty session.
2. In Word on the web: the document must not be read-only or locked by another user.
3. Open browser dev tools (F12) and check the Console tab for JavaScript errors from `officeInterop.js`.
4. Common error: `Word.run` fails with `RichAPI.Error: GeneralException` — usually means the document is protected. Ask the user to check **Review → Restrict Editing**.

---

## Task Pane Blank / JavaScript Error on Load

**Symptom:** The task pane opens but shows a spinner forever or a blank white page.

**Checklist:**
1. Open the task pane and press **Ctrl+Shift+I** (desktop) or right-click → **Inspect** to open DevTools.
2. Check Console for errors — common ones:
   - `office.js` failed to load → network/proxy blocking `appsforoffice.microsoft.com`
   - Blazor WASM failed to download → App Service may not be running; check App Service health
   - Mixed-content errors → the `SourceLocation` in manifest must be HTTPS
3. Verify the App Service is running:
   ```bash
   az webapp show \
     --name           REPLACE_APP_NAME \
     --resource-group REPLACE_RESOURCE_GROUP \
     --query          state
   ```

---

## SSL Certificate Errors (Local Dev)

**Symptom:** Browser blocks the add-in with "Your connection is not private" or cert warning.

**Fix:**
```powershell
dotnet dev-certs https --trust
```

After trusting, close and reopen Word (dev cert changes require a browser restart).

---

## Collecting Diagnostic Information

### Check the X-Auth-Error Header

The API adds an `X-Auth-Error` response header on all 401 responses. This contains the exact JWT validation error message (e.g. `IDX10214: Audience validation failed...`).

To read it:
1. Open browser DevTools (F12) → Network tab.
2. Find the failing API request (e.g. to `/api/snippets`).
3. Click on the request → Headers → Response Headers.
4. Look for `x-auth-error`.

### Decode a JWT

Paste any token at https://jwt.ms. Key claims to check:
- `aud` — must equal the Client ID GUID configured in `AzureAd:Audience`
- `iss` — must match `login.microsoftonline.com/{tenantId}/v2.0`
- `tid` — must match your Tenant ID
- `preferred_username` or `upn` — must be present (requires optional claims; see [Admin-Setup.md](Admin-Setup.md) step 3)
- `exp` — must be in the future (Unix timestamp)

### App Service Logs

```bash
az webapp log tail \
  --name           REPLACE_APP_NAME \
  --resource-group REPLACE_RESOURCE_GROUP
```

### Browser / Add-in Console Logs

1. Word Desktop: **Insert → Add-ins → My Add-ins** → right-click the add-in → **Inspect Element** (if available) or enable add-in debugging via registry.
2. Word Web: DevTools (F12) → Console tab.
3. Blazor WASM errors appear in the browser Console as structured messages.

### Azure SQL Audit Logs

If enabled, SQL audit logs are in Azure Monitor. Filter by connection failures:

```kusto
// Azure Monitor / Log Analytics
AzureDiagnostics
| where Category == "SQLSecurityAuditEvents"
| where action_name_s contains "FAILED_DATABASE_AUTHENTICATION"
| project TimeGenerated, client_ip_s, server_principal_name_s, statement_s
| order by TimeGenerated desc
```

---

## Quick Diagnostic Check Table

| Symptom | Most Likely Cause | Where to Look |
|---|---|---|
| Sign In button shown | Error 13006 (Word Online) — expected | Click Sign In; see MSAL Fallback section |
| 401 with `IDX10214` in X-Auth-Error | `Audience` in appsettings is the App ID URI instead of GUID | `appsettings.json` → `AzureAd:Audience` |
| 401 with `IDX10223` (issuer mismatch) | `TenantId` in appsettings is wrong | `appsettings.json` → `AzureAd:TenantId` |
| 401 with `IDX10501` (no token) | Token not attached to request | Check `officeInterop.js` `fetchJson` → is `Authorization` header set? |
| 500 on any request | SQL connection failure or missing optional claims | App Service logs; check `preferred_username` optional claim |
| Task pane blank | Blazor WASM download failed (App Service down/cold) | `az webapp show --query state` |
| Ribbon button missing | Centralized Deployment not propagated | M365 Admin Center → Integrated Apps |

---

## Getting Support

For issues not covered here:
1. Check the [Microsoft Office Add-in documentation](https://learn.microsoft.com/office/dev/add-ins/)
2. Post in [Microsoft Q&A](https://learn.microsoft.com/answers/tags/415/office-js)
3. File an issue in this repository with:
   - Office version (`File → Account → About Word`)
   - Browser (for Word on the web)
   - The exact error message and error code
   - The `X-Auth-Error` header value (if a 401)
   - Steps to reproduce
