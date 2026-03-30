# ============================================================================
# DWx Policy Manager - Master Provisioning Script
# Orchestrates the full deployment of Policy Manager to a SharePoint site
#
# Version: 1.2.5
# Date: 30 March 2026
# Company: First Digital
#
# USAGE:
#   .\00-Master-Provisioning.ps1
#   .\00-Master-Provisioning.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/PolicyManager"
#   .\00-Master-Provisioning.ps1 -SkipAzure -SkipSeedData
#
# PREREQUISITES:
#   - PnP.PowerShell module installed
#   - SharePoint Admin or Site Collection Admin access
#   - Azure CLI installed (if deploying Azure resources)
# ============================================================================

param(
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager",
    [switch]$SkipAzure,
    [switch]$SkipLists,
    [switch]$SkipConfig,
    [switch]$SkipSeedData,
    [switch]$SkipPages,
    [switch]$SkipVerification
)

$ErrorActionPreference = "Continue"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$provisioningRoot = Join-Path (Split-Path -Parent $scriptRoot) "policy-management"

# ============================================================================
# BANNER
# ============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  DWx Policy Manager - Master Provisioning" -ForegroundColor Cyan
Write-Host "  Version: 1.2.5" -ForegroundColor Cyan
Write-Host "  Target:  $SiteUrl" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Track overall results
$results = @{
    Passed = 0
    Failed = 0
    Skipped = 0
}
$stepErrors = @()

function Write-StepHeader {
    param([string]$StepNumber, [string]$StepName)
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  STEP $StepNumber: $StepName" -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Cyan
}

function Write-StepResult {
    param([string]$StepName, [bool]$Success, [string]$Message = "")
    if ($Success) {
        Write-Host "  [PASS] $StepName" -ForegroundColor Green
        $script:results.Passed++
    } else {
        Write-Host "  [FAIL] $StepName - $Message" -ForegroundColor Red
        $script:results.Failed++
        $script:stepErrors += "$StepName : $Message"
    }
}

function Write-StepSkipped {
    param([string]$StepName)
    Write-Host "  [SKIP] $StepName (skipped via parameter)" -ForegroundColor Yellow
    $script:results.Skipped++
}

# ============================================================================
# STEP 0: Validate Prerequisites
# ============================================================================

Write-StepHeader "0" "Validating Prerequisites"

# Check PnP PowerShell
$pnpModule = Get-Module -ListAvailable -Name "PnP.PowerShell"
if ($pnpModule) {
    Write-Host "  PnP.PowerShell v$($pnpModule.Version) found" -ForegroundColor Green
} else {
    Write-Host "  PnP.PowerShell not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
        Import-Module PnP.PowerShell -ErrorAction Stop
        Write-Host "  PnP.PowerShell installed successfully" -ForegroundColor Green
    } catch {
        Write-Host "  FATAL: Could not install PnP.PowerShell: $_" -ForegroundColor Red
        exit 1
    }
}

Import-Module PnP.PowerShell -ErrorAction Stop

# Check provisioning scripts directory
if (-not (Test-Path $provisioningRoot)) {
    Write-Host "  FATAL: Provisioning scripts not found at: $provisioningRoot" -ForegroundColor Red
    exit 1
}
Write-Host "  Provisioning scripts directory: $provisioningRoot" -ForegroundColor Green

# ============================================================================
# STEP 1: Connect to SharePoint
# ============================================================================

Write-StepHeader "1" "Connecting to SharePoint"

try {
    Write-Host "  Connecting to $SiteUrl ..." -ForegroundColor Cyan
    Write-Host "  A browser window will open for authentication." -ForegroundColor Yellow
    Connect-PnPOnline -Url $SiteUrl -Interactive
    $web = Get-PnPWeb -ErrorAction Stop
    Write-Host "  Connected to: $($web.Title) ($($web.Url))" -ForegroundColor Green
    Write-StepResult "SharePoint Connection" $true
} catch {
    Write-Host "  FATAL: Could not connect to SharePoint: $_" -ForegroundColor Red
    Write-StepResult "SharePoint Connection" $false $_.Exception.Message
    exit 1
}

# ============================================================================
# STEP 2: SharePoint List Provisioning
# ============================================================================

if ($SkipLists) {
    Write-StepSkipped "SharePoint List Provisioning"
} else {
    Write-StepHeader "2" "SharePoint List Provisioning"

    # Ordered list of provisioning scripts
    $listScripts = @(
        "01-Core-PolicyLists.ps1",
        "02-Quiz-Lists.ps1",
        "03-Exemption-Distribution-Lists.ps1",
        "04-Social-Lists.ps1",
        "05-PolicyPack-Lists.ps1",
        "06-Analytics-Audit-Lists.ps1",
        "07-Notification-Lists.ps1",
        "08-Approval-Lists.ps1",
        "09-PolicySourceDocuments.ps1",
        "10-CorporateTemplates.ps1",
        "11-AdminConfig-Lists.ps1",
        "12-UserManagement-Lists.ps1",
        "13-Visibility-Columns.ps1",
        "14-SubCategory-And-Folders.ps1",
        "15-ManagedDepartments-Column.ps1",
        "16-DistributionQueue-List.ps1",
        "16-TemplateType-Update.ps1",
        "17-ReportingLists.ps1",
        "18-MissingColumns-Patch.ps1",
        "19-PolicyRoleGroups.ps1",
        "20-NotificationChoiceUpdate.ps1",
        "21-EmailTemplates-List.ps1",
        "22-Audiences-List.ps1",
        "24-Distribution-Missing-Columns.ps1",
        "25-ReminderSchedule-List.ps1",
        "26-UserProfiles-Unified.ps1",
        "26-SLABreaches-List.ps1",
        "27-Social-Lists.ps1",
        "27-Missing-Lists-Master.ps1",
        "28-Retention-Lists.ps1",
        "29-Workflow-Lists.ps1"
    )

    $listSuccess = 0
    $listFail = 0

    foreach ($script in $listScripts) {
        $scriptPath = Join-Path $provisioningRoot $script
        if (Test-Path $scriptPath) {
            try {
                Write-Host "  Running: $script ..." -ForegroundColor Cyan
                & $scriptPath
                Write-Host "  [OK] $script" -ForegroundColor Green
                $listSuccess++
            } catch {
                Write-Host "  [ERROR] $script - $($_.Exception.Message)" -ForegroundColor Red
                $listFail++
            }
        } else {
            Write-Host "  [WARN] Script not found: $script" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "  List provisioning complete: $listSuccess succeeded, $listFail failed" -ForegroundColor $(if ($listFail -eq 0) { "Green" } else { "Yellow" })
    Write-StepResult "SharePoint List Provisioning" ($listFail -eq 0) "$listFail scripts failed"
}

# ============================================================================
# STEP 3: Configuration Seeding
# ============================================================================

if ($SkipConfig) {
    Write-StepSkipped "Configuration Seeding"
} else {
    Write-StepHeader "3" "Seeding PM_Configuration"

    $configEntries = @(
        @{ ConfigKey = "Integration.AI.Chat.Enabled";         ConfigValue = "true";           Category = "AI" },
        @{ ConfigKey = "Integration.AI.Chat.MaxTokens";       ConfigValue = "1000";           Category = "AI" },
        @{ ConfigKey = "Admin.Branding.CompanyName";          ConfigValue = "First Digital";  Category = "Branding" },
        @{ ConfigKey = "Admin.Branding.ProductName";          ConfigValue = "Policy Manager"; Category = "Branding" },
        @{ ConfigKey = "Admin.Upload.DocLimitMB";             ConfigValue = "25";             Category = "Admin" },
        @{ ConfigKey = "Admin.Upload.VideoLimitMB";           ConfigValue = "100";            Category = "Admin" },
        @{ ConfigKey = "Admin.Quiz.DefaultPassingScore";      ConfigValue = "70";             Category = "Quiz" },
        @{ ConfigKey = "Notifications.NewPolicy.Enabled";     ConfigValue = "true";           Category = "Notifications" },
        @{ ConfigKey = "Notifications.PolicyUpdate.Enabled";  ConfigValue = "true";           Category = "Notifications" },
        @{ ConfigKey = "Notifications.DailyDigest.Enabled";   ConfigValue = "false";          Category = "Notifications" },
        @{ ConfigKey = "Compliance.RequireAcknowledgement";   ConfigValue = "true";           Category = "Compliance" },
        @{ ConfigKey = "Compliance.DefaultDeadlineDays";      ConfigValue = "14";             Category = "Compliance" },
        @{ ConfigKey = "Compliance.ReviewFrequencyMonths";    ConfigValue = "12";             Category = "Compliance" },
        @{ ConfigKey = "Compliance.ReminderEnabled";          ConfigValue = "true";           Category = "Compliance" },
        @{ ConfigKey = "Approval.RequireOnNew";               ConfigValue = "true";           Category = "Approval" },
        @{ ConfigKey = "Approval.RequireOnUpdate";            ConfigValue = "true";           Category = "Approval" },
        @{ ConfigKey = "Approval.AllowSelfApproval";          ConfigValue = "false";          Category = "Approval" },
        @{ ConfigKey = "Theme.CustomEnabled";                 ConfigValue = "false";          Category = "Theme" },
        @{ ConfigKey = "Theme.PrimaryColor";                  ConfigValue = "#0d9488";        Category = "Theme" },
        @{ ConfigKey = "Theme.DarkColor";                     ConfigValue = "#0f766e";        Category = "Theme" }
    )

    $configList = "PM_Configuration"
    $configSuccess = 0
    $configSkipped = 0
    $configFailed = 0

    # Check if list exists
    $listExists = Get-PnPList -Identity $configList -ErrorAction SilentlyContinue
    if (-not $listExists) {
        Write-Host "  PM_Configuration list not found. Skipping config seeding." -ForegroundColor Yellow
        Write-Host "  Run list provisioning first (Step 2)." -ForegroundColor Yellow
        Write-StepResult "Configuration Seeding" $false "PM_Configuration list not found"
    } else {
        foreach ($entry in $configEntries) {
            try {
                # Check if key already exists
                $existing = Get-PnPListItem -List $configList -Query "<View><Query><Where><Eq><FieldRef Name='ConfigKey'/><Value Type='Text'>$($entry.ConfigKey)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
                if ($existing -and $existing.Count -gt 0) {
                    Write-Host "  [SKIP] $($entry.ConfigKey) (already exists)" -ForegroundColor Gray
                    $configSkipped++
                } else {
                    Add-PnPListItem -List $configList -Values @{
                        Title       = $entry.ConfigKey
                        ConfigKey   = $entry.ConfigKey
                        ConfigValue = $entry.ConfigValue
                        Category    = $entry.Category
                        IsActive    = "TRUE"
                    } | Out-Null
                    Write-Host "  [ADD] $($entry.ConfigKey) = $($entry.ConfigValue)" -ForegroundColor Green
                    $configSuccess++
                }
            } catch {
                Write-Host "  [ERROR] $($entry.ConfigKey) - $($_.Exception.Message)" -ForegroundColor Red
                $configFailed++
            }
        }

        Write-Host ""
        Write-Host "  Config seeding: $configSuccess added, $configSkipped existing, $configFailed failed" -ForegroundColor $(if ($configFailed -eq 0) { "Green" } else { "Yellow" })
        Write-StepResult "Configuration Seeding" ($configFailed -eq 0) "$configFailed entries failed"
    }
}

# ============================================================================
# STEP 4: Sample Data Seeding
# ============================================================================

if ($SkipSeedData) {
    Write-StepSkipped "Sample Data Seeding"
} else {
    Write-StepHeader "4" "Sample Data Seeding"

    $seedScripts = @(
        "Seed-CurrentUserData.ps1",
        "Seed-FastTrackTemplates.ps1",
        "Seed-ComprehensiveDemoData.ps1"
    )

    $seedSuccess = 0
    $seedFail = 0

    foreach ($script in $seedScripts) {
        $scriptPath = Join-Path $provisioningRoot $script
        if (Test-Path $scriptPath) {
            try {
                Write-Host "  Running: $script ..." -ForegroundColor Cyan
                & $scriptPath
                Write-Host "  [OK] $script" -ForegroundColor Green
                $seedSuccess++
            } catch {
                Write-Host "  [ERROR] $script - $($_.Exception.Message)" -ForegroundColor Red
                $seedFail++
            }
        } else {
            Write-Host "  [WARN] Script not found: $script" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "  Seed data: $seedSuccess succeeded, $seedFail failed" -ForegroundColor $(if ($seedFail -eq 0) { "Green" } else { "Yellow" })
    Write-StepResult "Sample Data Seeding" ($seedFail -eq 0) "$seedFail scripts failed"
}

# ============================================================================
# STEP 5: SharePoint Pages
# ============================================================================

if ($SkipPages) {
    Write-StepSkipped "SharePoint Pages"
} else {
    Write-StepHeader "5" "Creating SharePoint Pages"

    $pagesScript = Join-Path $provisioningRoot "Provision-SharePointPages.ps1"
    if (Test-Path $pagesScript) {
        try {
            & $pagesScript
            Write-StepResult "SharePoint Pages" $true
        } catch {
            Write-StepResult "SharePoint Pages" $false $_.Exception.Message
        }
    } else {
        Write-Host "  [WARN] Provision-SharePointPages.ps1 not found" -ForegroundColor Yellow
        Write-StepResult "SharePoint Pages" $false "Script not found"
    }

    # Create the 2 additional pages not in the original script
    $additionalPages = @(
        @{ Name = "PolicyAuthorReports"; Title = "Author Reports";     Description = "Author performance reports" },
        @{ Name = "PolicyBulkUpload";    Title = "Bulk Upload";        Description = "Bulk policy import with AI" }
    )

    foreach ($page in $additionalPages) {
        try {
            $existing = $null
            try { $existing = Get-PnPPage -Identity $page.Name -ErrorAction SilentlyContinue } catch { }
            if ($existing) {
                Write-Host "  [SKIP] $($page.Name).aspx already exists" -ForegroundColor Yellow
            } else {
                Add-PnPPage -Name $page.Name -Title $page.Title -LayoutType Article -CommentsEnabled:$false | Out-Null
                Write-Host "  [CREATE] $($page.Name).aspx - $($page.Description)" -ForegroundColor Green
            }
        } catch {
            Write-Host "  [ERROR] $($page.Name).aspx - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# ============================================================================
# STEP 6: Azure Infrastructure (Manual Reference)
# ============================================================================

if ($SkipAzure) {
    Write-StepSkipped "Azure Infrastructure"
} else {
    Write-StepHeader "6" "Azure Infrastructure"

    Write-Host ""
    Write-Host "  Azure resources must be deployed separately using the deploy.ps1" -ForegroundColor Yellow
    Write-Host "  scripts in each azure-functions/*/infra/ directory." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Required Azure deployments:" -ForegroundColor Cyan
    Write-Host "    1. Quiz Generator:    azure-functions/quiz-generator/infra/deploy.ps1" -ForegroundColor White
    Write-Host "    2. Chat Assistant:    azure-functions/policy-chat/infra/deploy.ps1" -ForegroundColor White
    Write-Host "    3. Email Sender:      azure-functions/email-sender/infra/deploy.ps1" -ForegroundColor White
    Write-Host "    4. Distribution:      azure-functions/distribution-processor/infra/deploy.ps1" -ForegroundColor White
    Write-Host "    5. Doc Converter:     azure-functions/document-converter/infra/deploy.ps1" -ForegroundColor White
    Write-Host "    6. Approval Escalation: azure-functions/approval-escalation/infra/deploy.ps1" -ForegroundColor White
    Write-Host ""
    Write-Host "  After deploying, update PM_Configuration with function URLs:" -ForegroundColor Yellow
    Write-Host "    - Integration.AI.Chat.FunctionUrl" -ForegroundColor White
    Write-Host "    - Integration.AI.Quiz.FunctionUrl" -ForegroundColor White
    Write-Host "    - Integration.DocConverter.FunctionUrl" -ForegroundColor White
    Write-Host ""
    Write-Host "  IMPORTANT: Authorize Logic App API connections in Azure Portal" -ForegroundColor Red
    Write-Host "    - office365-prod > Edit API connection > Authorize" -ForegroundColor White
    Write-Host "    - sharepointonline-prod > Edit API connection > Authorize" -ForegroundColor White
    Write-Host ""

    $results.Skipped++
}

# ============================================================================
# STEP 7: Verification
# ============================================================================

if ($SkipVerification) {
    Write-StepSkipped "Deployment Verification"
} else {
    Write-StepHeader "7" "Deployment Verification"

    $verifyScript = Join-Path $scriptRoot "07-Verify-Deployment.ps1"
    if (Test-Path $verifyScript) {
        try {
            & $verifyScript
            Write-StepResult "Deployment Verification" $true
        } catch {
            Write-StepResult "Deployment Verification" $false $_.Exception.Message
        }
    } else {
        Write-Host "  [WARN] 07-Verify-Deployment.ps1 not found at: $verifyScript" -ForegroundColor Yellow
        Write-Host "  Skipping automated verification." -ForegroundColor Yellow
        $results.Skipped++
    }
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  DEPLOYMENT SUMMARY" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Target:   $SiteUrl" -ForegroundColor White
Write-Host "  Passed:   $($results.Passed)" -ForegroundColor Green
Write-Host "  Failed:   $($results.Failed)" -ForegroundColor $(if ($results.Failed -gt 0) { "Red" } else { "Green" })
Write-Host "  Skipped:  $($results.Skipped)" -ForegroundColor Yellow
Write-Host ""

if ($stepErrors.Count -gt 0) {
    Write-Host "  ERRORS:" -ForegroundColor Red
    foreach ($err in $stepErrors) {
        Write-Host "    - $err" -ForegroundColor Red
    }
    Write-Host ""
}

if ($results.Failed -eq 0) {
    Write-Host "  Deployment completed successfully!" -ForegroundColor Green
} else {
    Write-Host "  Deployment completed with $($results.Failed) error(s)." -ForegroundColor Yellow
    Write-Host "  Review the errors above and re-run failed steps." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  NEXT STEPS:" -ForegroundColor Cyan
Write-Host "    1. Add webparts to each SharePoint page (see Deployment Guide Section 4.4)" -ForegroundColor White
Write-Host "    2. Deploy Azure resources if not done (see Section 3)" -ForegroundColor White
Write-Host "    3. Update PM_Configuration with Azure function URLs" -ForegroundColor White
Write-Host "    4. Import users from CSV or sync from Entra ID" -ForegroundColor White
Write-Host "    5. Run manual testing checklist (see Deployment Guide Section 9.2)" -ForegroundColor White
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
