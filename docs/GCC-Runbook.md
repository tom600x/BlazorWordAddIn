# GCC Runbook

Guidance specific to deploying and maintaining the BlazorWordAddIn in a **US Government Community Cloud (GCC)** Microsoft 365 tenant backed by commercial Azure.

> **Scope:** This runbook targets **GCC only** (not GCC High, not DoD).  
> GCC High and DoD tenants use `login.microsoftonline.us` and Azure Government infrastructure — see variant notes at the end of this document.

---

## 1. GCC vs GCC High vs DoD — Quick Reference

| Attribute | GCC | GCC High | DoD |
|---|---|---|---|
| AAD Endpoint | `login.microsoftonline.com` | `login.microsoftonline.us` | `login.microsoftonline.us` |
| Azure subscription | Commercial or Azure Gov | Azure Government only | Azure Government DoD only |
| Office CDN | `appsforoffice.microsoft.com` | `appsforoffice.microsoft.com` | `appsforoffice.microsoft.com` |
| Add-in store | Not available | Not available | Not available |
| Deployment method | Centralized Deployment (admin portal) | Centralized Deployment | Centralized Deployment |
| JWT `iss` claim | `login.microsoftonline.com/{tenantId}/v2.0` | `login.microsoftonline.us/{tenantId}/v2.0` | `login.microsoftonline.us/{tenantId}/v2.0` |
| JWT `aud` claim | Client ID GUID | Client ID GUID | Client ID GUID |

> **Authentication note for GCC:** GCC tenants authenticate through **commercial Azure AD** (`login.microsoftonline.com`). This is the default configuration of this add-in. No changes to `AzureAd:Instance` are needed for GCC.

---

## 2. Network Requirements

Ensure the following endpoints are reachable from client machines (not blocked by proxy/firewall):

| Endpoint | Purpose |
|---|---|
| `https://appsforoffice.microsoft.com` | Office.js CDN loaded by the add-in |
| `https://login.microsoftonline.com` | Azure AD authentication (GCC uses commercial endpoint) |
| `https://graph.microsoft.com` | Microsoft Graph (for profile claims) |
| `https://YOUR_APP_NAME.azurewebsites.net` | Add-in UI and backend API |
| `*.database.windows.net` | Azure SQL (if using public endpoint) |

If the environment uses a proxy, configure `HTTPS_PROXY` or organization proxy PAC files accordingly.

---

## 3. GCC Authentication Configuration

The `appsettings.json` for GCC is identical to the standard commercial configuration:

```json
"AzureAd": {
  "Instance":  "https://login.microsoftonline.com/",
  "TenantId":  "REPLACE_TENANT_GUID",
  "ClientId":  "REPLACE_ADDIN_CLIENT_GUID",
  "Audience":  "REPLACE_ADDIN_CLIENT_GUID"
}
```

> **Audience must be the Client ID GUID.** Do not set `Audience` to the App ID URI (`api://HOSTNAME/GUID`). MSAL tokens — whether from Office SSO or the `auth.html` MSAL fallback — set `aud = CLIENT_GUID`. The full App ID URI is only used for scoping in the Office manifest's `WebApplicationInfo`, not for token audience validation. Setting `Audience` to the App ID URI causes `IDX10214` errors on every request.

---

## 4. MSAL Fallback Behavior in GCC Word Online

In Word Online (browser-based), `Office.auth.getAccessToken()` returns error **13006** because the Office client runs in a sandboxed iframe. This is **expected behavior in GCC environments** — it is not a misconfiguration.

The add-in handles this through an MSAL dialog fallback:

1. The task pane detects error 13006.
2. A **Sign In** button is displayed.
3. User clicks **Sign In** → the Office Dialog API opens `auth.html` in a popup.
4. `auth.html` runs MSAL `loginRedirect` through `login.microsoftonline.com`.
5. After sign-in, `auth.html` posts the JWT back to the task pane.

**GCC-specific note:** Ensure `https://YOUR_APP_NAME.azurewebsites.net/auth.html` is registered as a **Single-page application** redirect URI (not a Web redirect URI) in your Entra app registration. The `app-registration.ps1` script sets this automatically. If done manually, see [Admin-Setup.md](Admin-Setup.md) step 1g.

---

## 5. Centralized Deployment Steps

Centralized Deployment pushes the add-in to targeted users/groups without requiring action from each user.

### 5.1 Prerequisites

- Exchange Online deployed in the GCC tenant
- Users have Exchange Online mailboxes
- Global Admin or Exchange Admin role
- Microsoft 365 Apps for Government (formerly ProPlus) license minimum

### 5.2 Deploy the Add-in

1. Navigate to [https://admin.microsoft.com](https://admin.microsoft.com).
2. In the left navigation, go to **Settings → Integrated apps**.
3. Click **Upload custom apps**.
4. Under **App type**, select **Office Add-in**.
5. Select **Upload manifest file (.xml) from device** and upload `add-in/manifest.xml`.
6. Click **Next**.
7. Select **Assign users**:
   - Choose **Specific users / groups** for a pilot rollout (recommended for first deployment).
   - Choose **Entire organization** for full rollout.
8. Click **Next**, review the permissions list, check the consent box, and click **Finish deployment**.
9. Click **Done**. The add-in status shows **Deployed** after propagation (15–30 minutes).

### 5.3 Verify Deployment

After propagation:
1. **Close all Office applications** on a test machine.
2. Relaunch Microsoft Word.
3. Click the **Home** tab.
4. Look for the **Word Snippets** button in the ribbon.
5. Alternatively, go to **Insert → Add-ins → Admin Managed** — the add-in should appear there.

### 5.4 Update an Existing Deployment

1. Admin Center → **Integrated Apps** → find the add-in by name.
2. Click **Update**.
3. Upload the new `add-in/manifest.xml`.
4. Confirm the update. Propagation typically takes 1–2 hours; no user action required.

---

## 6. Office Client Minimum Builds (GCC)

Office SSO (`Office.auth.getAccessToken`) requires modern Office builds. Users on older builds will always hit error 13006 and must use the MSAL fallback Sign In button.

| Platform | Minimum build for SSO |
|---|---|
| Microsoft 365 Apps for Windows | 16.0.14326.20454 |
| Microsoft 365 for Mac | 16.55 |
| Word on the web | All current builds (SSO unsupported — MSAL fallback used) |
| Word for iPad | 2.72 |
| Word for Android | 16.0.14326 |

To check build: **File → Account → About Word** (desktop) or tenant admin page.  
Push updates via **Servicing Profiles** in M365 admin if users are on Semi-Annual Channel.

---

## 7. SSO Consent in GCC

Office SSO requires one-time **admin consent** to grant the add-in access to Microsoft Graph scopes (`openid`, `profile`, `User.Read`).

### CLI

```bash
az ad app permission admin-consent --id REPLACE_ADDIN_CLIENT_ID
```

### Portal

1. Go to [https://portal.azure.com](https://portal.azure.com).
2. Navigate to **Entra ID → App registrations → BlazorWordAddIn → API permissions**.
3. Click **Grant admin consent for \<tenant name\>**.
4. Click **Yes** to confirm.

Without admin consent, individual users receive a consent prompt the first time they open the add-in.

---

## 8. Azure Infrastructure Options for GCC Customers

GCC tenants can deploy their Azure backend to either:

### Option A: Commercial Azure (Default)

- Region examples: `eastus`, `westus2`, `centralus`
- SQL Server DNS: `*.database.windows.net`
- App Service DNS: `*.azurewebsites.net`
- **No code changes needed** — this is the default configuration.

### Option B: Azure Government Cloud

For GCC customers who require data residency within the FedRAMP High boundary:

| Parameter | Commercial (Option A) | Azure Government (Option B) |
|---|---|---|
| `location` | `eastus` | `usgovvirginia` |
| SQL endpoint suffix | `.database.windows.net` | `.database.usgovcloudapi.net` |
| App Service DNS | `*.azurewebsites.net` | `*.azurewebsites.us` |
| AAD endpoint (`AzureAd:Instance`) | `login.microsoftonline.com` | `login.microsoftonline.com` ← still commercial! |
| Azure CLI environment | default | `az cloud set --name AzureUSGovernment` |

> **Critical:** Even when deploying infrastructure to Azure Government, **GCC tenants still authenticate through commercial AAD** (`login.microsoftonline.com`). Do **not** change `AzureAd:Instance` to `login.microsoftonline.us`. Only GCC High/DoD tenants use the `.us` endpoint.

To deploy to Azure Government:

```bash
az cloud set --name AzureUSGovernment
az login
# Deploy as normal using the infra scripts
```

Update `api/appsettings.json` connection string to use the `.database.usgovcloudapi.net` suffix for Option B.

---

## 9. Data Classification Banner

For environments that require classification markings, the add-in's task pane header can display a banner. Add the following to the App Service application settings (Azure portal → App Service → Configuration → Application settings):

| Name | Value |
|---|---|
| `AppSettings__DataClassificationBanner` | `UNCLASSIFIED // FOR OFFICIAL USE ONLY` |

The Blazor UI reads this value from the API configuration endpoint and displays it at the top of the task pane.

---

## 10. Security Notes

- The add-in uses Office SSO or MSAL dialog — no separate username/password is required or stored.
- All tokens are short-lived JWTs (1 hour default) validated by the backend on every request.
- Azure SQL uses Managed Identity — no SQL passwords are stored, transmitted, or rotated.
- All traffic is TLS 1.2+; HTTP is redirected to HTTPS at the App Service level.
- Each user can only see their own snippets, scoped by their UPN.
- The add-in does not write to Word documents without explicit user action (clicking the **Insert** button).

---

## 11. Incident Response — Add-in Not Working (GCC Helpdesk)

Quick triage for GCC helpdesk staff:

| Symptom | First Action | Escalation |
|---|---|---|
| Ribbon button missing | Check Centralized Deployment status in M365 Admin Center → Integrated Apps | Verify user is in the deployment target group; wait 30 min for propagation |
| Task pane blank / loading forever | Check network access to `appsforoffice.microsoft.com` from the user's machine | Check App Service health in Azure portal |
| "Sign In" button shown (Word Online) | This is expected — user must click Sign In | If dialog doesn't open, check corporate pop-up blocker settings |
| "Unable to get access token" | Verify user is signed in with their work account (File → Account → Sign in) | Check SSO error code; see [Troubleshooting.md](Troubleshooting.md) error code table |
| 401 errors after sign-in | Check `AzureAd:Audience` in App Service app settings — must be the Client GUID, not `api://...` | Decode JWT at jwt.ms; check `aud` claim |
| Data not loading (500 errors) | Verify App Service is running; check App Service logs | Check SQL Managed Identity and firewall settings |
| Wrong user's data shown | SSO token UPN claim mismatch | Verify optional claims are configured in Entra; see [Admin-Setup.md](Admin-Setup.md) step 3 |

See [Troubleshooting.md](Troubleshooting.md) for full error code reference.
