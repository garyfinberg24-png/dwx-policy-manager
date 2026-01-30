# ============================================================================
# DWx Policy Manager — Quiz Generator Infrastructure Deployment
# ============================================================================
# Prerequisites:
#   - Azure CLI installed (az --version)
#   - Logged in to Azure (az login)
#   - Bicep CLI installed (az bicep install)
#   - Azure Functions Core Tools (npm i -g azure-functions-core-tools@4)
#
# Usage:
#   .\deploy.ps1                          # Deploy dev (default)
#   .\deploy.ps1 -Environment prod        # Deploy prod
#   .\deploy.ps1 -SkipInfra               # Skip infra, deploy code only
#   .\deploy.ps1 -SkipCode                # Deploy infra only
# ============================================================================

param(
    [ValidateSet("dev", "staging", "prod")]
    [string]$Environment = "dev",

    [string]$Location = "australiaeast",

    [string]$BaseName = "dwx-pm",

    [string]$SubscriptionName = "",

    [switch]$SkipInfra,

    [switch]$SkipCode,

    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"
$InfraDir = $PSScriptRoot
$FunctionDir = Split-Path $InfraDir -Parent

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
# Validation
# ============================================================================

Write-Step "Validating prerequisites"

# Check Azure CLI
try {
    $azVersion = az version --output json | ConvertFrom-Json
    Write-Info "Azure CLI: $($azVersion.'azure-cli')"
} catch {
    Write-Error "Azure CLI is not installed. Install from https://aka.ms/installazurecli"
}

# Check Bicep
try {
    $bicepVersion = az bicep version 2>&1
    Write-Info "Bicep: $bicepVersion"
} catch {
    Write-Warn "Bicep not found. Installing..."
    az bicep install
}

# Check logged in
$account = az account show --output json 2>$null | ConvertFrom-Json
if (-not $account) {
    Write-Error "Not logged in to Azure. Run: az login"
}
Write-Info "Logged in as: $($account.user.name)"
Write-Info "Subscription: $($account.name)"

# Switch subscription if specified
if ($SubscriptionName) {
    Write-Info "Switching to subscription: $SubscriptionName"
    az account set --subscription $SubscriptionName
}

# ============================================================================
# Variables
# ============================================================================

$ResourceGroupName = "$BaseName-quiz-rg-$Environment"
$FunctionAppName = "$BaseName-quiz-func-$Environment"

Write-Step "Deployment Configuration"
Write-Info "Environment:     $Environment"
Write-Info "Location:        $Location"
Write-Info "Resource Group:  $ResourceGroupName"
Write-Info "Function App:    $FunctionAppName"
Write-Info "Base Name:       $BaseName"
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
            --tags project="DWx Policy Manager" component="Quiz Generator" environment=$Environment `
            --output none

        Write-Success "Resource group '$ResourceGroupName' ready"
    }

    # ============================================================================
    # Step 2: Validate Bicep Template
    # ============================================================================

    Write-Step "Step 2/4 — Validating Bicep Template"

    $templateFile = Join-Path $InfraDir "main.bicep"
    $parametersFile = Join-Path $InfraDir "main.parameters.json"

    if (-not (Test-Path $templateFile)) {
        Write-Error "Bicep template not found: $templateFile"
    }

    # Override parameters for this deployment
    $overrides = @(
        "baseName=$BaseName",
        "location=$Location",
        "environment=$Environment"
    )

    Write-Info "Validating template..."
    az deployment group validate `
        --resource-group $ResourceGroupName `
        --template-file $templateFile `
        --parameters $parametersFile `
        --parameters $overrides `
        --output none

    Write-Success "Template validation passed"

    # ============================================================================
    # Step 3: Deploy Bicep Template
    # ============================================================================

    Write-Step "Step 3/4 — Deploying Infrastructure"

    if ($WhatIf) {
        Write-Info "[WHAT-IF] Running what-if analysis..."
        az deployment group what-if `
            --resource-group $ResourceGroupName `
            --template-file $templateFile `
            --parameters $parametersFile `
            --parameters $overrides
    } else {
        Write-Info "Deploying resources (this may take several minutes)..."

        $deploymentOutput = az deployment group create `
            --resource-group $ResourceGroupName `
            --template-file $templateFile `
            --parameters $parametersFile `
            --parameters $overrides `
            --name "quiz-generator-$Environment-$(Get-Date -Format 'yyyyMMdd-HHmmss')" `
            --output json | ConvertFrom-Json

        if ($LASTEXITCODE -ne 0) {
            Write-Error "Infrastructure deployment failed"
        }

        # Extract outputs
        $outputs = $deploymentOutput.properties.outputs
        $functionUrl = $outputs.functionAppUrl.value
        $openAiEndpoint = $outputs.openAiEndpoint.value
        $kvName = $outputs.keyVaultName.value

        Write-Success "Infrastructure deployed"
        Write-Info "Function App URL: $functionUrl"
        Write-Info "OpenAI Endpoint:  $openAiEndpoint"
        Write-Info "Key Vault:        $kvName"
    }
} else {
    Write-Step "Skipping infrastructure deployment (--SkipInfra)"
}

# ============================================================================
# Step 4: Deploy Function Code
# ============================================================================

if (-not $SkipCode) {
    Write-Step "Step 4/4 — Deploying Function Code"

    if ($WhatIf) {
        Write-Info "[WHAT-IF] Would build and deploy function code to $FunctionAppName"
    } else {
        # Build TypeScript
        Write-Info "Building TypeScript..."
        Push-Location $FunctionDir
        try {
            npm install --silent 2>$null
            npm run build
            if ($LASTEXITCODE -ne 0) {
                Write-Error "TypeScript build failed"
            }
            Write-Success "TypeScript build complete"

            # Deploy to Azure
            Write-Info "Publishing to Azure Functions..."
            func azure functionapp publish $FunctionAppName --typescript
            if ($LASTEXITCODE -ne 0) {
                Write-Error "Function deployment failed"
            }
            Write-Success "Function code deployed to $FunctionAppName"
        } finally {
            Pop-Location
        }

        # Get function key for SPFx configuration
        Write-Info "Retrieving function key..."
        $funcKeys = az functionapp keys list `
            --resource-group $ResourceGroupName `
            --name $FunctionAppName `
            --output json | ConvertFrom-Json

        $hostKey = $funcKeys.functionKeys.default
        if ($hostKey) {
            Write-Host ""
            Write-Host "  +-------------------------------------------------+" -ForegroundColor Green
            Write-Host "  |  DEPLOYMENT COMPLETE                             |" -ForegroundColor Green
            Write-Host "  +-------------------------------------------------+" -ForegroundColor Green
            Write-Host ""
            Write-Host "  Function URL:" -ForegroundColor White
            Write-Host "  https://$FunctionAppName.azurewebsites.net/api/generate-quiz-questions" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Function Key (add to QuizBuilder config):" -ForegroundColor White
            Write-Host "  $hostKey" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Full URL with key:" -ForegroundColor White
            Write-Host "  https://$FunctionAppName.azurewebsites.net/api/generate-quiz-questions?code=$hostKey" -ForegroundColor Yellow
            Write-Host ""
        } else {
            Write-Warn "Could not retrieve function key. Get it from Azure Portal:"
            Write-Warn "  Portal > $FunctionAppName > Functions > generateQuizQuestions > Function Keys"
        }
    }
} else {
    Write-Step "Skipping code deployment (--SkipCode)"
}

# ============================================================================
# Summary
# ============================================================================

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " Deployment Summary" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Resource Group:  $ResourceGroupName" -ForegroundColor White
Write-Host "  Function App:    $FunctionAppName" -ForegroundColor White
Write-Host "  Environment:     $Environment" -ForegroundColor White
Write-Host "  Location:        $Location" -ForegroundColor White
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor Gray
Write-Host "  1. Copy the Function URL + key into the QuizBuilder AI panel" -ForegroundColor Gray
Write-Host "  2. Test: POST to /api/generate-quiz-questions with a sample policy" -ForegroundColor Gray
Write-Host "  3. Monitor: Azure Portal > $FunctionAppName > Monitor" -ForegroundColor Gray
Write-Host ""
