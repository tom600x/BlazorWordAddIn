// ─── Azure SQL Server + Database ─────────────────────────────────────────────
// Authentication: Azure AD (Entra ID) only. No SQL admin password is stored.
// The App Service's managed identity is granted access via post-deploy script.

@description('Azure region.')
param location string

@description('SQL Server name (globally unique).')
param serverName string

@description('SQL Database name.')
param databaseName string

// ─── SQL Server ───────────────────────────────────────────────────────────────
resource sqlServer 'Microsoft.Sql/servers@2023-05-01-preview' = {
  name: serverName
  location: location
  properties: {
    // Azure AD only authentication – no SQL login passwords stored.
    // An Entra admin must be configured post-deploy via CLI (see post-deploy.ps1).
    administrators: {
      administratorType: 'ActiveDirectory'
      azureADOnlyAuthentication: false  // set true after configuring Entra admin
      login: 'REPLACE_ENTRA_ADMIN_DISPLAY_NAME'
      sid: 'REPLACE_ENTRA_ADMIN_OBJECT_ID'
      tenantId: 'REPLACE_TENANT_ID'
    }
    minimalTlsVersion: '1.2'
    publicNetworkAccess: 'Enabled'  // required to allow Azure services
  }
}

// ─── Firewall: allow traffic from Azure services ──────────────────────────────
resource allowAzureServices 'Microsoft.Sql/servers/firewallRules@2023-05-01-preview' = {
  parent: sqlServer
  name: 'AllowAllWindowsAzureIps'
  properties: {
    startIpAddress: '0.0.0.0'
    endIpAddress: '0.0.0.0'
  }
}

// ─── SQL Database ─────────────────────────────────────────────────────────────
resource sqlDatabase 'Microsoft.Sql/servers/databases@2023-05-01-preview' = {
  parent: sqlServer
  name: databaseName
  location: location
  sku: {
    name: 'S0'    // Standard S0 (10 DTUs) – upgrade as needed
    tier: 'Standard'
  }
  properties: {
    collation: 'SQL_Latin1_General_CP1_CI_AS'
  }
}

// ─── Outputs ─────────────────────────────────────────────────────────────────
output serverFqdn string = sqlServer.properties.fullyQualifiedDomainName
