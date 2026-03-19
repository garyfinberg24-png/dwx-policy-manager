# ============================================================================
# DWx Policy Manager — Document Converter Deployment
# ============================================================================
# Deploys an Azure Function that converts Office documents (.docx, .pptx, .xlsx)
# to clean HTML for the Policy Manager reader.
#
# Prerequisites:
#   - Azure CLI installed (az --version)
#   - Logged in to Azure (az login)
#   - Bicep CLI installed (az bicep install)
#   - App Registration with SharePoint Sites.Read.All (for downloading files)
#
# Usage:
#   .\deploy.ps1                                    # Deploy prod (default)
#   .\deploy.ps1 -Environment dev                   # Deploy dev
#   .\deploy.ps1 -WhatIf                            # Dry run
#   .\deploy.ps1 -SkipInfra                         # Skip infra, deploy code only
#
# Post-deployment:
#   1. Retrieve the function key from Azure Portal
#   2. Configure the full URL (with ?code=) in Policy Manager Admin > AI Settings
# ============================================================================

param(
    [ValidateSet("dev", "staging", "prod")]
    [string]$Environment = "prod",

    [string]$Location = "australiaeast",

    [string]$BaseName = "dwx-pm",

    [string]$SubscriptionName = "",

    [switch]$SkipInfra,

    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"
$InfraDir = $PSScriptRoot
$ProjectDir = Split-Path $InfraDir -Parent

# ============================================================================
# Helpers
# ============================================================================

function Write-Step($message) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " $message" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
}

function Write-Info($message) {
    Write-Host "  [INFO] $message" -ForegroundColor Gray
}

function Write-Success($message) {
    Write-Host "  [OK]   $message" -ForegroundColor Green
}

function Write-Warn($message) {
    Write-Host "  [WARN] $message" -ForegroundColor Yellow
}

# ============================================================================
# Subscription (optional override)
# ============================================================================

if ($SubscriptionName) {
    Write-Info "Switching to subscription: $SubscriptionName"
    az account set --subscription $SubscriptionName
}

# ============================================================================
# Variables
# ============================================================================

$ResourceGroupName = "$BaseName-docconv-rg-$Environment"
$FunctionAppName = "$BaseName-docconv-func-$Environment"
$TemplateFile = Join-Path $InfraDir "main.bicep"

Write-Step "Deployment Configuration"
Write-Info "Environment:      $Environment"
Write-Info "Location:         $Location"
Write-Info "Resource Group:   $ResourceGroupName"
Write-Info "Function App:     $FunctionAppName"
Write-Host ""

# ============================================================================
# Step 1: Create Resource Group
# ============================================================================

if (-not $SkipInfra) {
    Write-Step "Step 1/4 — Creating Resource Group"

    if ($WhatIf) {
        Write-Info "[WHAT-IF] Would create resource group: $ResourceGroupName in $Location"
    } else {
        az group create `
            --name $ResourceGroupName `
            --location $Location `
            --tags project="DWx Policy Manager" component="Document Converter" environment=$Environment `
            --output none

        Write-Success "Resource group '$ResourceGroupName' ready"
    }

    # ============================================================================
    # Step 2: Validate Bicep Template
    # ============================================================================

    Write-Step "Step 2/4 — Validating Bicep Template"

    if ($WhatIf) {
        Write-Info "[WHAT-IF] Would validate Bicep template"
    } else {
        az deployment group validate `
            --resource-group $ResourceGroupName `
            --template-file $TemplateFile `
            --parameters baseName=$BaseName location=$Location environment=$Environment `
            --output none

        Write-Success "Template validation passed"
    }

    # ============================================================================
    # Step 3: Deploy Bicep Template
    # ============================================================================

    Write-Step "Step 3/4 — Deploying Infrastructure"

    if ($WhatIf) {
        Write-Info "[WHAT-IF] Running what-if analysis..."
        az deployment group what-if `
            --resource-group $ResourceGroupName `
            --template-file $TemplateFile `
            --parameters baseName=$BaseName location=$Location environment=$Environment
    } else {
        Write-Info "Deploying resources (this may take a minute)..."

        $deploymentOutput = az deployment group create `
            --resource-group $ResourceGroupName `
            --template-file $TemplateFile `
            --parameters baseName=$BaseName location=$Location environment=$Environment `
            --name "docconv-func-$Environment-$(Get-Date -Format 'yyyyMMdd-HHmmss')" `
            --output json | ConvertFrom-Json

        if ($LASTEXITCODE -ne 0) {
            Write-Error "Deployment failed"
        }

        Write-Success "Infrastructure deployed"
    }
} else {
    Write-Info "Skipping infrastructure deployment (-SkipInfra)"
}

# ============================================================================
# Step 4: Deploy Function Code
# ============================================================================

Write-Step "Step 4/4 — Deploying Function Code"

if ($WhatIf) {
    Write-Info "[WHAT-IF] Would build and deploy function code"
} else {
    Write-Info "Installing dependencies..."
    Push-Location $ProjectDir
    npm install --production 2>&1 | Out-Null

    Write-Info "Building TypeScript..."
    npm run build 2>&1 | Out-Null

    Write-Info "Publishing to Azure..."
    func azure functionapp publish $FunctionAppName --typescript 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Warn "func publish failed. Trying zip deploy fallback..."

        # Zip deploy fallback
        $zipPath = Join-Path $env:TEMP "docconv-func.zip"
        Write-Info "Creating deployment zip..."

        # Create zip with required files
        $filesToZip = @(
            "dist",
            "node_modules",
            "host.json",
            "package.json"
        )

        if (Test-Path $zipPath) { Remove-Item $zipPath }
        Compress-Archive -Path ($filesToZip | ForEach-Object { Join-Path $ProjectDir $_ }) -DestinationPath $zipPath -Force

        Write-Info "Deploying via zip..."
        az functionapp deployment source config-zip `
            --resource-group $ResourceGroupName `
            --name $FunctionAppName `
            --src $zipPath `
            --output none

        if ($LASTEXITCODE -eq 0) {
            Write-Success "Function code deployed (zip deploy)"
        } else {
            Write-Error "Both func publish and zip deploy failed"
        }

        Remove-Item $zipPath -ErrorAction SilentlyContinue
    } else {
        Write-Success "Function code deployed"
    }

    Pop-Location

    # Retrieve function key
    Write-Info "Retrieving function key..."
    $keys = az functionapp keys list --name $FunctionAppName --resource-group $ResourceGroupName --output json 2>$null | ConvertFrom-Json
    $funcKey = $keys.functionKeys.default

    if ($funcKey) {
        $fullUrl = "https://$FunctionAppName.azurewebsites.net/api/convertDocument?code=$funcKey"
        Write-Success "Function key retrieved"
    } else {
        $fullUrl = "https://$FunctionAppName.azurewebsites.net/api/convertDocument?code=<retrieve-from-portal>"
        Write-Warn "Could not retrieve function key automatically. Get it from Azure Portal."
    }
}

# ============================================================================
# Summary
# ============================================================================

Write-Host ""
Write-Host "  +-------------------------------------------------+" -ForegroundColor Green
Write-Host "  |  DEPLOYMENT COMPLETE                             |" -ForegroundColor Green
Write-Host "  +-------------------------------------------------+" -ForegroundColor Green
Write-Host ""
Write-Host "  Function App:   $FunctionAppName" -ForegroundColor White
Write-Host "  Resource Group: $ResourceGroupName" -ForegroundColor White
Write-Host "  Environment:    $Environment" -ForegroundColor White
Write-Host ""
Write-Host "  NEXT STEPS:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. Copy the Function URL below" -ForegroundColor Yellow
Write-Host "  2. Go to Policy Manager > Admin > AI Settings" -ForegroundColor Yellow
Write-Host "  3. Paste the URL into 'Document Converter Function URL'" -ForegroundColor Yellow
Write-Host "  4. Click 'Test Connection' to verify" -ForegroundColor Yellow
Write-Host ""
if (-not $WhatIf) {
    Write-Host "  Function URL:" -ForegroundColor Cyan
    Write-Host "  $fullUrl" -ForegroundColor White
}
Write-Host ""
