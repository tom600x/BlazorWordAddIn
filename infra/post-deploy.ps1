<#
.SYNOPSIS
    Grants the App Service's Managed Identity database access on Azure SQL.

.DESCRIPTION
    After deploying the infrastructure, the App Service has a system-assigned
    Managed Identity, but it cannot connect to SQL yet.  This script:
        1. Obtains the App Service's MI display name
        2. Connects to the database as the signed-in user (who must be the Entra
           SQL admin or have sufficient rights)
        3. Creates a contained database user mapped to the MI
        4. Assigns db_datareader and db_datawriter roles

    IMPORTANT: You must run this connected to the database as an Entra user with
    db_owner or equivalent permissions.  Run 'az login' before executing.

.NOTES
    Prerequisites:
        - SqlServer PowerShell module:  Install-Module -Name SqlServer -Force
        - Azure CLI installed and logged in
        - You are the Entra admin of the SQL Server

.PARAMETER ResourceGroup
    Azure resource group name.

.PARAMETER AppName
    App Service name (also the Managed Identity display name).

.PARAMETER SqlServerFqdn
    Fully qualified SQL server hostname, e.g. myserver.database.windows.net

.PARAMETER DatabaseName
    Target database name (default: SnippetsDb).

.EXAMPLE
    .\post-deploy.ps1 -ResourceGroup rg-word-snippets -AppName word-snippets-prod `
                      -SqlServerFqdn myserver.database.windows.net -DatabaseName SnippetsDb
#>
param(
    [Parameter(Mandatory = $true)]  [string]$ResourceGroup,
    [Parameter(Mandatory = $true)]  [string]$AppName,
    [Parameter(Mandatory = $true)]  [string]$SqlServerFqdn,
    [Parameter(Mandatory = $false)] [string]$DatabaseName = 'SnippetsDb'
)

$ErrorActionPreference = 'Stop'

# ── Verify SqlServer module ───────────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name SqlServer)) {
    Write-Host "Installing SqlServer PowerShell module..." -ForegroundColor Cyan
    Install-Module -Name SqlServer -Force -AllowClobber -Scope CurrentUser
}
Import-Module SqlServer

# ── Get an access token for the current CLI user ─────────────────────────────
Write-Host "Acquiring Azure SQL access token for current CLI user..." -ForegroundColor Cyan
$tokenJson = az account get-access-token --resource https://database.windows.net/ | ConvertFrom-Json
$accessToken = $tokenJson.accessToken

# ── Run SQL to grant the Managed Identity access ─────────────────────────────
$sql = @"
-- Create contained user mapped to the App Service Managed Identity
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = '$AppName')
BEGIN
    CREATE USER [$AppName] FROM EXTERNAL PROVIDER;
END

-- Grant read/write access
ALTER ROLE db_datareader ADD MEMBER [$AppName];
ALTER ROLE db_datawriter ADD MEMBER [$AppName];

PRINT 'Managed Identity ''$AppName'' granted db_datareader + db_datawriter on $DatabaseName.';
"@

Write-Host "Connecting to $SqlServerFqdn / $DatabaseName ..." -ForegroundColor Cyan
Invoke-Sqlcmd `
    -ServerInstance $SqlServerFqdn `
    -Database $DatabaseName `
    -AccessToken $accessToken `
    -Query $sql

Write-Host "`n✅ Managed Identity access granted." -ForegroundColor Green
Write-Host "   App Service '$AppName' can now connect to '$DatabaseName' via Managed Identity."
