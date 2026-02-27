# ============================================================================
# DWx Policy Manager — Email Queue Processor (Logic App) Deployment
# ============================================================================
# Deploys an Azure Logic App that:
#   1. Polls PM_EmailQueue SharePoint list every N minutes
#   2. Sends queued emails via Office 365 connector
#   3. Updates queue item status (Sent/Failed) with retry logic
#
# Prerequisites:
#   - Azure CLI installed (az --version)
#   - Logged in to Azure (az login)
#   - Bicep CLI installed (az bicep install)
#   - Access to the target SharePoint site
#
# Usage:
#   .\deploy.ps1                                    # Deploy prod (default)
#   .\deploy.ps1 -Environment dev                   # Deploy dev
#   .\deploy.ps1 -WhatIf                            # Dry run
#   .\deploy.ps1 -PollingMinutes 2                  # Poll every 2 minutes
#   .\deploy.ps1 -SenderEmail "noreply@company.com" # Use shared mailbox
#
# Post-deployment:
#   After deploying, you MUST authorize the API connections in the Azure Portal:
#   1. Go to Resource Group → dwx-pm-email-rg-{env}
#   2. Click on 'office365-{env}' connection → Edit API connection → Authorize
#   3. Click on 'sharepointonline-{env}' connection → Edit API connection → Authorize
#   4. Run the Logic App manually once to verify
# ============================================================================

param(
    [ValidateSet("dev", "staging", "prod")]
    [string]$Environment = "prod",

    [string]$Location = "australiaeast",

    [string]$BaseName = "dwx-pm",

    [string]$SubscriptionName = "",

    [int]$PollingMinutes = 5,

    [int]$BatchSize = 20,

    [string]$SenderEmail = "",

    [string]$SharePointSiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager",

    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"
$InfraDir = $PSScriptRoot

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

$ResourceGroupName = "$BaseName-email-rg-$Environment"
$LogicAppName = "$BaseName-email-sender-$Environment"
$TemplateFile = Join-Path $InfraDir "main.bicep"
$ParametersFile = Join-Path $InfraDir "main.parameters.json"

Write-Step "Deployment Configuration"
Write-Info "Environment:      $Environment"
Write-Info "Location:         $Location"
Write-Info "Resource Group:   $ResourceGroupName"
Write-Info "Logic App:        $LogicAppName"
Write-Info "SP Site:          $SharePointSiteUrl"
Write-Info "Polling Interval: Every $PollingMinutes minutes"
Write-Info "Batch Size:       $BatchSize emails per run"
if ($SenderEmail) {
    Write-Info "Sender Email:     $SenderEmail"
} else {
    Write-Info "Sender Email:     (connection user's mailbox)"
}
Write-Host ""

# Validate template exists
if (-not (Test-Path $TemplateFile)) {
    Write-Error "Bicep template not found: $TemplateFile"
}

# ============================================================================
# Step 1: Create Resource Group
# ============================================================================

Write-Step "Step 1/3 — Creating Resource Group"

if ($WhatIf) {
    Write-Info "[WHAT-IF] Would create resource group: $ResourceGroupName in $Location"
} else {
    az group create `
        --name $ResourceGroupName `
        --location $Location `
        --tags project="DWx Policy Manager" component="Email Sender" environment=$Environment `
        --output none

    Write-Success "Resource group '$ResourceGroupName' ready"
}

# ============================================================================
# Step 2: Validate Bicep Template
# ============================================================================

Write-Step "Step 2/3 — Validating Bicep Template"

# Override parameters for this deployment
$overrides = @(
    "baseName=$BaseName",
    "location=$Location",
    "environment=$Environment",
    "sharePointSiteUrl=$SharePointSiteUrl",
    "pollingIntervalMinutes=$PollingMinutes",
    "batchSize=$BatchSize",
    "maxRetryAttempts=3",
    "senderEmailAddress=$SenderEmail"
)

Write-Info "Validating template..."

if ($WhatIf) {
    Write-Info "[WHAT-IF] Would validate Bicep template"
} else {
    az deployment group validate `
        --resource-group $ResourceGroupName `
        --template-file $TemplateFile `
        --parameters $ParametersFile `
        --parameters $overrides `
        --output none

    Write-Success "Template validation passed"
}

# ============================================================================
# Step 3: Deploy Bicep Template
# ============================================================================

Write-Step "Step 3/3 — Deploying Logic App"

if ($WhatIf) {
    Write-Info "[WHAT-IF] Running what-if analysis..."
    az deployment group what-if `
        --resource-group $ResourceGroupName `
        --template-file $TemplateFile `
        --parameters $ParametersFile `
        --parameters $overrides
} else {
    Write-Info "Deploying resources (this may take a minute)..."

    $deploymentOutput = az deployment group create `
        --resource-group $ResourceGroupName `
        --template-file $TemplateFile `
        --parameters $ParametersFile `
        --parameters $overrides `
        --name "email-sender-$Environment-$(Get-Date -Format 'yyyyMMdd-HHmmss')" `
        --output json | ConvertFrom-Json

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Deployment failed"
    }

    # Extract outputs
    $outputs = $deploymentOutput.properties.outputs

    Write-Success "Logic App deployed"
    Write-Info "Logic App:       $($outputs.logicAppName.value)"
    Write-Info "O365 Connection: $($outputs.office365ConnectionId.value)"
    Write-Info "SP Connection:   $($outputs.sharepointConnectionId.value)"
}

# ============================================================================
# Summary
# ============================================================================

Write-Host ""
Write-Host "  +-------------------------------------------------+" -ForegroundColor Green
Write-Host "  |  DEPLOYMENT COMPLETE                             |" -ForegroundColor Green
Write-Host "  +-------------------------------------------------+" -ForegroundColor Green
Write-Host ""
Write-Host "  Logic App:       $LogicAppName" -ForegroundColor White
Write-Host "  Resource Group:  $ResourceGroupName" -ForegroundColor White
Write-Host "  Environment:     $Environment" -ForegroundColor White
Write-Host "  Polling:         Every $PollingMinutes minutes" -ForegroundColor White
Write-Host ""
Write-Host "  IMPORTANT — You must authorize the API connections:" -ForegroundColor Yellow
Write-Host ""
Write-Host "  1. Go to Azure Portal > Resource Group: $ResourceGroupName" -ForegroundColor Yellow
Write-Host "  2. Click 'office365-$Environment' > Edit API connection > Authorize" -ForegroundColor Yellow
Write-Host "     Sign in with a mailbox that has permission to send emails" -ForegroundColor Yellow
Write-Host "  3. Click 'sharepointonline-$Environment' > Edit API connection > Authorize" -ForegroundColor Yellow
Write-Host "     Sign in with an account that has access to PM_EmailQueue" -ForegroundColor Yellow
Write-Host "  4. Open the Logic App > Overview > Run Trigger > Run" -ForegroundColor Yellow
Write-Host "     Verify the first run succeeds in the Run History tab" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Email Flow:" -ForegroundColor Cyan
Write-Host "  SPFx App --> EmailQueueService --> PM_EmailQueue (SP List)" -ForegroundColor Gray
Write-Host "           --> Logic App polls every $PollingMinutes min" -ForegroundColor Gray
Write-Host "           --> Office 365 sends email" -ForegroundColor Gray
Write-Host "           --> Status updated: Queued -> Processing -> Sent/Failed" -ForegroundColor Gray
Write-Host "           --> Failed emails retry up to 3 times" -ForegroundColor Gray
Write-Host ""
