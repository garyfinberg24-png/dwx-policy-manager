# ============================================================
# Deploy Distribution Queue Processor Azure Function
# ============================================================
param(
    [string]$Environment = "prod",
    [string]$Location = "australiaeast",
    [string]$TenantId = "",
    [string]$ClientId = "",
    [string]$ClientSecret = ""
)

$ErrorActionPreference = "Stop"
$rgName = "dwx-pm-dist-rg-$Environment"

Write-Host "`n=== Deploying Distribution Queue Processor ===" -ForegroundColor Cyan
Write-Host "Resource Group: $rgName" -ForegroundColor Gray
Write-Host "Location: $Location" -ForegroundColor Gray

# Create resource group
az group create --name $rgName --location $Location --output none

# Deploy Bicep template
az deployment group create `
    --resource-group $rgName `
    --template-file "$PSScriptRoot/main.bicep" `
    --parameters `
        environment=$Environment `
        location=$Location `
        tenantId=$TenantId `
        clientId=$ClientId `
        clientSecret=$ClientSecret `
    --output none

# Get function app name
$funcAppName = (az deployment group show `
    --resource-group $rgName `
    --name main `
    --query "properties.outputs.functionAppName.value" `
    --output tsv)

Write-Host "`nFunction App: $funcAppName" -ForegroundColor Green

# Deploy function code
Write-Host "`nBuilding and deploying function code..." -ForegroundColor Yellow
Push-Location "$PSScriptRoot/.."
npm install
npm run build

# Create deployment zip
$zipPath = "$env:TEMP\dist-processor.zip"
Compress-Archive -Path "dist/*", "host.json", "package.json", "node_modules/*" -DestinationPath $zipPath -Force

az functionapp deployment source config-zip `
    --resource-group $rgName `
    --name $funcAppName `
    --src $zipPath

Pop-Location

Write-Host "`n=== Deployment Complete ===" -ForegroundColor Cyan
Write-Host "Function App: https://$funcAppName.azurewebsites.net" -ForegroundColor Green
Write-Host "Timer: Runs every 2 minutes" -ForegroundColor Green
Write-Host "`nRequired: Azure AD App Registration with Sites.FullControl.All permission" -ForegroundColor Yellow
