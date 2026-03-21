targetScope = 'resourceGroup'

// ─── Parameters ──────────────────────────────────────────────────────────────
@description('Azure region for all resources.')
param location string = resourceGroup().location

@description('Base name used for App Service, SQL Server, etc. (3-24 chars, lowercase alphanumeric).')
@minLength(3)
@maxLength(24)
param appName string

@description('Azure SQL Server name (must be globally unique).')
param sqlServerName string

@description('Azure SQL Database name.')
param sqlDatabaseName string = 'SnippetsDb'

@description('Entra application (client) ID of the Word add-in app registration.')
param addinClientId string

@description('Entra tenant ID.')
param tenantId string

// ─── Module: App Service ─────────────────────────────────────────────────────
module appService 'modules/appservice.bicep' = {
  name: 'appService'
  params: {
    location: location
    appName: appName
    sqlServerName: sqlServerName
    sqlDatabaseName: sqlDatabaseName
    addinClientId: addinClientId
    tenantId: tenantId
  }
}

// ─── Module: Azure SQL ───────────────────────────────────────────────────────
module sql 'modules/sql.bicep' = {
  name: 'sql'
  params: {
    location: location
    serverName: sqlServerName
    databaseName: sqlDatabaseName
  }
}

// ─── Outputs ─────────────────────────────────────────────────────────────────
@description('Default hostname of the deployed App Service. Use this as REPLACE_APP_HOSTNAME.')
output appHostname string = appService.outputs.appHostname

@description('Full HTTPS URL of the app.')
output appUrl string = 'https://${appService.outputs.appHostname}'
