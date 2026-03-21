<#
.SYNOPSIS
    Builds the .NET solution and deploys it to Azure App Service via Bicep.

.DESCRIPTION
    1. Reads parameters from infra/parameters.bicepparam
    2. Creates the resource group if it does not exist
    3. Deploys main.bicep
    4. Publishes the ASP.NET Core (hosted Blazor WASM) app to App Service
    5. Prints next steps for post-deploy SQL configuration

.NOTES
    Prerequisites:
        - Azure CLI installed and logged in:  az login --tenant <TENANT_ID>
        - .NET 8 SDK installed
        - Resource group already exists, or this script creates it

    Usage:
        .\deploy.ps1 -ResourceGroup rg-word-snippets -Location eastus
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroup,

    [Parameter(Mandatory = $false)]
    [string]$Location = 'eastus',

    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId  # optional: if omitted uses current az context subscription
)

$ErrorActionPreference = 'Stop'
$ScriptDir = $PSScriptRoot
$RepoRoot  = Split-Path -Parent $ScriptDir

# ── 0. Optionally set subscription ───────────────────────────────────────────
if ($SubscriptionId) {
    Write-Host "Setting subscription: $SubscriptionId" -ForegroundColor Cyan
    az account set --subscription $SubscriptionId
}

# ── 1. Create resource group if needed ───────────────────────────────────────
$rgExists = az group exists --name $ResourceGroup
if ($rgExists -eq 'false') {
    Write-Host "Creating resource group '$ResourceGroup' in '$Location'..." -ForegroundColor Cyan
    az group create --name $ResourceGroup --location $Location | Out-Null
}

# ── 2. Deploy Bicep ──────────────────────────────────────────────────────────
Write-Host "`nDeploying infrastructure (Bicep)..." -ForegroundColor Cyan

$deploymentOutput = az deployment group create `
    --resource-group $ResourceGroup `
    --template-file "$ScriptDir\main.bicep" `
    --parameters "$ScriptDir\parameters.bicepparam" `
    --output json | ConvertFrom-Json

if ($LASTEXITCODE -ne 0) {
    Write-Error "Bicep deployment failed."
}

$appHostname = $deploymentOutput.properties.outputs.appHostname.value
$appUrl      = $deploymentOutput.properties.outputs.appUrl.value

Write-Host "  App Hostname : $appHostname" -ForegroundColor Green
Write-Host "  App URL      : $appUrl"      -ForegroundColor Green

# ── 3. Build and publish the .NET application ────────────────────────────────
Write-Host "`nPublishing .NET solution..." -ForegroundColor Cyan

$publishDir = Join-Path $ScriptDir ".publish"
if (Test-Path $publishDir) { Remove-Item $publishDir -Recurse -Force }

dotnet publish "$RepoRoot\api\Api.csproj" `
    --configuration Release `
    --output $publishDir `
    --nologo

if ($LASTEXITCODE -ne 0) { Write-Error "dotnet publish failed." }

# ── 4. Deploy to Azure App Service ───────────────────────────────────────────
Write-Host "`nDeploying app to Azure App Service..." -ForegroundColor Cyan

# Create a zip of the publish output
$zipPath = Join-Path $ScriptDir "app-publish.zip"
Compress-Archive -Path "$publishDir\*" -DestinationPath $zipPath -Force

# Read app name from parameters file (quick parse)
$paramsContent = Get-Content "$ScriptDir\parameters.bicepparam" -Raw
$appNameMatch  = [regex]::Match($paramsContent, "param appName\s+=\s+'([^']+)'")
$appName       = $appNameMatch.Groups[1].Value

az webapp deploy `
    --resource-group $ResourceGroup `
    --name $appName `
    --src-path $zipPath `
    --type zip

if ($LASTEXITCODE -ne 0) { Write-Error "App Service deployment failed." }

# ── Cleanup ───────────────────────────────────────────────────────────────────
Remove-Item $zipPath -Force
Remove-Item $publishDir -Recurse -Force

# ── 5. Print next steps ───────────────────────────────────────────────────────
Write-Host "`n✅ Deployment complete!" -ForegroundColor Green
Write-Host "
NEXT STEPS
──────────
1. Run post-deploy.ps1 to grant Managed Identity access to SQL:
   .\post-deploy.ps1 -ResourceGroup $ResourceGroup -AppName $appName

2. Run app-registration.ps1 to configure the Entra app registration (if not done):
   .\app-registration.ps1 -TenantId <TENANT_ID> -AppHostname $appHostname

3. Update infra\parameters.bicepparam and add-in\manifest.xml with:
   - REPLACE_APP_HOSTNAME → $appHostname
   - REPLACE_ADDIN_CLIENT_ID → <your Entra client ID>

4. Deploy the manifest to Microsoft 365 admin center (Centralized Deployment).
   See docs\GCC-Runbook.md for GCC-specific steps.

App URL: $appUrl
"
