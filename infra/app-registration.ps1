<#
.SYNOPSIS
    Creates and configures the Entra (Azure AD) app registration for the Word add-in.

.DESCRIPTION
    Steps performed:
        1. Create the app registration
        2. Set the Application ID URI  (api://HOSTNAME/CLIENT_ID)
        3. Add the 'access_as_user' delegated scope
        4. Pre-authorize the well-known Office client IDs to enable SSO without prompt
        5. Add required Microsoft Graph delegated permissions (openid, profile, User.Read)
        6. Print the client ID and instructions for adding it to the manifest

.NOTES
    Prerequisites:
        - Azure CLI installed and logged in as a user with 'Application Administrator'
          or 'Cloud Application Administrator' role in Entra
        - jq installed (optional; used for cleaner JSON parsing), OR PowerShell ConvertFrom-Json

.PARAMETER TenantId
    Entra tenant ID.

.PARAMETER AppHostname
    Hostname where the add-in is hosted, e.g. word-snippets.azurewebsites.net
    Used to set the Application ID URI and redirect URI.

.PARAMETER AppDisplayName
    Display name for the app registration (default: BlazorWordAddIn).

.EXAMPLE
    .\app-registration.ps1 `
        -TenantId 00000000-0000-0000-0000-000000000000 `
        -AppHostname word-snippets.azurewebsites.net
#>
param(
    [Parameter(Mandatory = $true)]  [string]$TenantId,
    [Parameter(Mandatory = $true)]  [string]$AppHostname,
    [Parameter(Mandatory = $false)] [string]$AppDisplayName = 'BlazorWordAddIn'
)

$ErrorActionPreference = 'Stop'

Write-Host "Logging in to tenant $TenantId..." -ForegroundColor Cyan
az login --tenant $TenantId --allow-no-subscriptions --scope "https://graph.microsoft.com//.default" | Out-Null

# Helper: serialise a PSObject to a temp JSON file and return the path.
# az rest / az ad app update accept  --body @filepath  on all platforms.
function Write-TempJson ($obj) {
    $f = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), '.json')
    # [System.IO.File]::WriteAllText writes UTF-8 without BOM on both PS 5.1 and PS 7.
    # Set-Content -Encoding utf8NoBOM only exists in PS 7+; avoid it for portability.
    $json = $obj | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($f, $json, [System.Text.Encoding]::UTF8)
    return $f
}

# ── 1. Create or reuse the app registration ───────────────────────────────────
Write-Host "Looking up app registration '$AppDisplayName'..." -ForegroundColor Cyan

$existing = az ad app list --display-name $AppDisplayName | ConvertFrom-Json
if ($existing.Count -gt 0) {
    $appJson = $existing[0]
    Write-Host "  Reusing existing app registration." -ForegroundColor Yellow
} else {
    Write-Host "Creating app registration '$AppDisplayName'..." -ForegroundColor Cyan
    # Create with web redirect URI first; SPA redirect URI is added via PATCH below
    # because --spa-redirect-uris is not available in all az ad app create versions.
    $appJson = az ad app create `
        --display-name $AppDisplayName `
        --sign-in-audience AzureADMyOrg `
        --enable-access-token-issuance $true `
        --enable-id-token-issuance $true `
        | ConvertFrom-Json
}

$ClientId    = $appJson.appId
$AppObjectId = $appJson.id

Write-Host "  Client ID   : $ClientId" -ForegroundColor Green
Write-Host "  Object ID   : $AppObjectId"

# ── 2. Set Application ID URI and SPA redirect URIs ─────────────────────────
# SPA redirect URIs are required for MSAL auth.html fallback (used when Office
# SSO returns error 13006, which is common in Word Online).
Write-Host "Setting Application ID URI and SPA redirect URIs..." -ForegroundColor Cyan

az ad app update `
    --id $ClientId `
    --identifier-uris "api://$AppHostname/$ClientId"

# Add SPA redirect URIs via Graph PATCH (az ad app update does not support --spa-redirect-uris on all versions)
$spaPayload = @{
    spa = @{
        redirectUris = @(
            "https://$AppHostname/auth.html"
        )
    }
}
$spaFile = Write-TempJson $spaPayload
az rest `
    --method PATCH `
    --uri "https://graph.microsoft.com/v1.0/applications/$AppObjectId" `
    --headers 'Content-Type=application/json' `
    --body "@$spaFile"
Remove-Item $spaFile -ErrorAction SilentlyContinue
Write-Host "  SPA redirect URI set: https://$AppHostname/auth.html" -ForegroundColor Gray

# ── 3. Add 'access_as_user' scope + pre-authorize Office clients ──────────────
# Scope and pre-auth are combined into a single PATCH to avoid a race condition
# where the pre-auth references a scope ID that isn't committed yet.
Write-Host "Adding 'access_as_user' scope and pre-authorizing Office clients..." -ForegroundColor Cyan

# Reuse the existing scope ID if the scope was already committed; otherwise generate a new one.
# Generating a fresh UUID on every run causes the pre-auth to reference an ID that doesn't match
# what's stored in the app, producing an "InvalidValue" error on subsequent runs.
$currentApp  = az ad app show --id $ClientId | ConvertFrom-Json
$existingScope = $currentApp.api.oauth2PermissionScopes | Where-Object { $_.value -eq 'access_as_user' }
$scopeId = if ($existingScope) { $existingScope.id } else { [System.Guid]::NewGuid().ToString() }

$officeClientIds = @(
    'd3590ed6-52b3-4102-aeff-aad2292ab01c'
    'ea5a67f6-b6f3-4338-b240-c655ddc3cc8e'
    '93d53678-613d-4013-afc3-0ebc4b444d30'
    'bc59ab01-8403-45c6-8796-ac3ef710b3e3'
    '57fb890c-0dab-4253-a5e0-7188c88b2bb4'
    '08e18876-6177-487e-b8b5-cf950c1e598c'
)

# Graph validates each pre-authorized appId against service principals in the tenant.
# Filter to only the IDs that are actually provisioned here.
Write-Host "  Checking which Office client IDs exist in this tenant..." -ForegroundColor Gray
$validOfficeClientIds = $officeClientIds | Where-Object {
    $sp = az ad sp list --filter "appId eq '$_'" 2>$null | ConvertFrom-Json
    $sp.Count -gt 0
}
Write-Host "  Found $($validOfficeClientIds.Count) of $($officeClientIds.Count) Office client SPs in this tenant." -ForegroundColor Gray

# Step A: commit the scope first; Graph validates pre-auth refs against already-stored scopes.
$scopePayload = @{
    api = @{
        oauth2PermissionScopes = @(
            @{
                id                      = $scopeId
                type                    = 'User'
                value                   = 'access_as_user'
                adminConsentDisplayName = 'Access Word Snippets on behalf of the user'
                adminConsentDescription = 'Allows the add-in to access the backend API on behalf of the signed-in user.'
                userConsentDisplayName  = 'Access Word Snippets'
                userConsentDescription  = 'Allows the add-in to access the Word Snippets service on your behalf.'
                isEnabled               = $true
            }
        )
    }
}

$scopeFile = Write-TempJson $scopePayload
az rest `
    --method PATCH `
    --uri "https://graph.microsoft.com/v1.0/applications/$AppObjectId" `
    --headers 'Content-Type=application/json' `
    --body "@$scopeFile"
Remove-Item $scopeFile -ErrorAction SilentlyContinue

# Step B: pre-authorize only the Office clients whose SPs exist in this tenant.
if ($validOfficeClientIds.Count -gt 0) {
    $preAuthPayload = @{
        api = @{
            # @() forces an array even when only one SP was found; without it PowerShell
            # returns a plain hashtable and ConvertTo-Json emits {} instead of [{}],
            # which fails Graph's schema validation.
            preAuthorizedApplications = @($validOfficeClientIds | ForEach-Object {
                @{
                    appId                  = $_
                    delegatedPermissionIds = @($scopeId)
                }
            })
        }
    }

    $preAuthFile = Write-TempJson $preAuthPayload
    az rest `
        --method PATCH `
        --uri "https://graph.microsoft.com/v1.0/applications/$AppObjectId" `
        --headers 'Content-Type=application/json' `
        --body "@$preAuthFile"
    Remove-Item $preAuthFile -ErrorAction SilentlyContinue
} else {
    Write-Warning "No Office client service principals found in this tenant. Pre-authorization skipped."
    Write-Warning "Add pre-authorized apps manually via Entra portal: App registrations -> Expose an API -> Add a client application."
}

# ── 4. Add required Graph delegated permissions ───────────────────────────────
# --required-resource-accesses expects a JSON array written to a file (@path).
# Using --set requiredResourceAccess= causes a schema-validation error.
Write-Host "Adding Microsoft Graph delegated permissions..." -ForegroundColor Cyan

$reqResPayload = @{
    requiredResourceAccess = @(
        @{
            resourceAppId  = '00000003-0000-0000-c000-000000000000'
            resourceAccess = @(
                @{ id = '37f7f235-527c-4136-accd-4a02d197296e'; type = 'Scope' }
                @{ id = '14dad69e-099b-42c9-810b-d002981feec1'; type = 'Scope' }
                @{ id = 'e1fe6dd8-ba31-4d61-89e7-88639da4683d'; type = 'Scope' }
            )
        }
    )
}

$reqResFile = Write-TempJson $reqResPayload
az rest `
    --method PATCH `
    --uri "https://graph.microsoft.com/v1.0/applications/$AppObjectId" `
    --headers 'Content-Type=application/json' `
    --body "@$reqResFile"
Remove-Item $reqResFile -ErrorAction SilentlyContinue

# ── 5. Add optional claims ───────────────────────────────────────────────────
# Optional claims ensure the JWT contains preferred_username, upn, email, and oid.
# Without these the backend cannot determine the user's UPN and returns 401/empty list.
Write-Host "Adding optional claims (preferred_username, upn, email, oid)..." -ForegroundColor Cyan

$optionalClaimsPayload = @{
    optionalClaims = @{
        accessToken = @(
            @{ name = 'upn';                  essential = $false }
            @{ name = 'email';                essential = $false }
            @{ name = 'preferred_username';   essential = $false }
            @{ name = 'oid';                  essential = $false }
        )
        idToken = @(
            @{ name = 'upn';                  essential = $false }
            @{ name = 'email';                essential = $false }
            @{ name = 'preferred_username';   essential = $false }
            @{ name = 'oid';                  essential = $false }
        )
    }
}
$claimsFile = Write-TempJson $optionalClaimsPayload
az rest `
    --method PATCH `
    --uri "https://graph.microsoft.com/v1.0/applications/$AppObjectId" `
    --headers 'Content-Type=application/json' `
    --body "@$claimsFile"
Remove-Item $claimsFile -ErrorAction SilentlyContinue
Write-Host "  Optional claims configured." -ForegroundColor Gray

# ── Print summary ─────────────────────────────────────────────────────────────
Write-Host @"

✅ App Registration complete!

  Display Name        : $AppDisplayName
  Client ID           : $ClientId
  Application ID URI  : api://$AppHostname/$ClientId
  access_as_user ID   : $scopeId
  SPA Redirect URI    : https://$AppHostname/auth.html

ACTION REQUIRED  – update the following files with the values above:
  1. add-in\manifest.xml
       <Id>$ClientId</Id>
       <Resource>api://$AppHostname/$ClientId</Resource>
  2. infra\parameters.bicepparam
       param addinClientId = '$ClientId'
  3. api\appsettings.json  (also via App Service config after Bicep deploy)
       "ClientId"  : "$ClientId"
       "Audience"  : "$ClientId"   <-- NOTE: use the GUID, NOT the full App ID URI.
                                       MSAL SPA tokens set aud = client GUID.

NOTE: Admin consent for the Graph scopes must be granted in the Azure portal:
  Entra → App Registrations → $AppDisplayName → API Permissions → Grant admin consent

"@ -ForegroundColor Green
