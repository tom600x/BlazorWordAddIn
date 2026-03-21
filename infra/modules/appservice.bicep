// ─── App Service Plan + Web App ───────────────────────────────────────────────

@description('Azure region.')
param location string

@description('Name for the App Service (also becomes the default hostname).')
param appName string

@description('SQL Server name for connection string config.')
param sqlServerName string

@description('SQL Database name for connection string config.')
param sqlDatabaseName string

@description('Entra application (client) ID.')
param addinClientId string

@description('Entra tenant ID.')
param tenantId string

// ─── App Service Plan ────────────────────────────────────────────────────────
resource appServicePlan 'Microsoft.Web/serverfarms@2023-01-01' = {
  name: '${appName}-plan'
  location: location
  sku: {
    name: 'B1'    // B1 minimum for custom domain + managed cert; upgrade for production
    tier: 'Basic'
  }
  kind: 'linux'
  properties: {
    reserved: true  // required for Linux
  }
}

// ─── Web App ─────────────────────────────────────────────────────────────────
resource webApp 'Microsoft.Web/sites@2023-01-01' = {
  name: appName
  location: location
  kind: 'app,linux'
  identity: {
    type: 'SystemAssigned'  // Managed Identity used to connect to SQL
  }
  properties: {
    serverFarmId: appServicePlan.id
    httpsOnly: true
    siteConfig: {
      linuxFxVersion: 'DOTNETCORE|8.0'
      minTlsVersion: '1.2'
      http20Enabled: true
      appSettings: [
        {
          name: 'ASPNETCORE_ENVIRONMENT'
          value: 'Production'
        }
        {
          name: 'AzureAd__Instance'
          value: 'https://login.microsoftonline.com/'
          // NOTE for GCC: GCC (not GCC-High) tenants use the commercial
          // login.microsoftonline.com endpoint – NOT login.microsoftonline.us.
        }
        {
          name: 'AzureAd__TenantId'
          value: tenantId
        }
        {
          name: 'AzureAd__ClientId'
          value: addinClientId
        }
        {
          name: 'AzureAd__Audience'
          // MSAL SPA tokens (acquired via PublicClientApplication / loginRedirect) set
          // aud = the client GUID, NOT the full App ID URI.  The backend must match
          // what the token actually contains, which is just the client ID GUID.
          value: addinClientId
        }
        {
          name: 'ConnectionStrings__DefaultConnection'
          // Authentication=Active Directory Default → resolved via DefaultAzureCredential
          // (uses the system-assigned Managed Identity on App Service in production)
          value: 'Server=tcp:${sqlServerName}.database.windows.net,1433;Initial Catalog=${sqlDatabaseName};Authentication=Active Directory Default;TrustServerCertificate=False;Encrypt=True;'
        }
        {
          name: 'Cors__AllowedOrigins__0'
          value: 'https://${appName}.azurewebsites.net'
        }
      ]
    }
  }
}

// ─── Outputs ─────────────────────────────────────────────────────────────────
output appHostname string = webApp.properties.defaultHostName
output principalId string = webApp.identity.principalId
