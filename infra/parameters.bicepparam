using 'main.bicep'

// ── Required parameters ────────────────────────────────────────────────────────
// Replace every REPLACE_ value before running deploy.ps1.

param appName           = 'REPLACE_APP_NAME'           // e.g. 'word-snippets-prod'
param sqlServerName     = 'REPLACE_SQL_SERVER_NAME'    // globally unique SQL server name
param sqlDatabaseName   = 'SnippetsDb'
param addinClientId     = 'c7806dca-7fab-4e78-bf17-d3f977b3eea4'    // Entra app registration client ID
param tenantId          = '17784004-2abc-4e6e-ade2-374b35b35643'          // Entra tenant ID
param location          = 'eastus'                     // Azure region
