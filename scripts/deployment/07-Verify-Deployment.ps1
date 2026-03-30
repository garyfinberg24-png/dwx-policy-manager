# ============================================================================
# DWx Policy Manager - Deployment Verification Script
# Validates that all required SharePoint components are provisioned correctly
#
# Version: 1.2.5
# Date: 30 March 2026
# Company: First Digital
#
# USAGE:
#   Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive
#   .\07-Verify-Deployment.ps1
#
# PREREQUISITES:
#   - Already connected to SharePoint via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  DWx Policy Manager - Deployment Verification" -ForegroundColor Cyan
Write-Host "  Version: 1.2.5" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Verify connection
try {
    $web = Get-PnPWeb -ErrorAction Stop
    Write-Host "  Connected to: $($web.Title) ($($web.Url))" -ForegroundColor Green
} catch {
    Write-Host "  ERROR: Not connected to SharePoint!" -ForegroundColor Red
    Write-Host "  Please connect first:" -ForegroundColor Yellow
    Write-Host '  Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive' -ForegroundColor White
    exit 1
}

# ============================================================================
# Tracking
# ============================================================================

$totalChecks = 0
$passed = 0
$failed = 0
$warnings = 0
$failedItems = @()

function Test-Check {
    param(
        [string]$Category,
        [string]$Name,
        [bool]$Result,
        [string]$Detail = ""
    )
    $script:totalChecks++
    if ($Result) {
        Write-Host "  [PASS] $Name" -ForegroundColor Green
        if ($Detail) { Write-Host "         $Detail" -ForegroundColor Gray }
        $script:passed++
    } else {
        Write-Host "  [FAIL] $Name" -ForegroundColor Red
        if ($Detail) { Write-Host "         $Detail" -ForegroundColor Red }
        $script:failed++
        $script:failedItems += "$Category : $Name"
    }
}

function Test-Warning {
    param(
        [string]$Name,
        [string]$Detail = ""
    )
    Write-Host "  [WARN] $Name" -ForegroundColor Yellow
    if ($Detail) { Write-Host "         $Detail" -ForegroundColor Yellow }
    $script:warnings++
}

# ============================================================================
# CHECK 1: Core SharePoint Lists
# ============================================================================

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  CHECK 1: Core SharePoint Lists" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$requiredLists = @(
    # Core Policy
    @{ Name = "PM_Policies";                Category = "Core" },
    @{ Name = "PM_PolicyVersions";          Category = "Core" },
    @{ Name = "PM_PolicyAcknowledgements";  Category = "Core" },
    @{ Name = "PM_PolicyMetadataProfiles";  Category = "Core" },
    @{ Name = "PM_PolicyReviewers";         Category = "Core" },
    @{ Name = "PM_PolicyCategories";        Category = "Core" },
    @{ Name = "PM_PolicyRequests";          Category = "Core" },
    @{ Name = "PM_PolicySubCategories";     Category = "Core" },
    @{ Name = "PM_PolicySourceDocuments";   Category = "Core" },
    @{ Name = "PM_PolicyTemplates";         Category = "Core" },
    @{ Name = "PM_PolicyExemptions";        Category = "Core" },
    @{ Name = "PM_PolicyDistributions";     Category = "Core" },
    @{ Name = "PM_DistributionQueue";       Category = "Core" },

    # Quiz
    @{ Name = "PM_PolicyQuizzes";           Category = "Quiz" },
    @{ Name = "PM_PolicyQuizQuestions";     Category = "Quiz" },
    @{ Name = "PM_PolicyQuizResults";       Category = "Quiz" },

    # Approval
    @{ Name = "PM_Approvals";              Category = "Approval" },
    @{ Name = "PM_ApprovalHistory";        Category = "Approval" },
    @{ Name = "PM_ApprovalDelegations";    Category = "Approval" },

    # Notification
    @{ Name = "PM_Notifications";          Category = "Notification" },
    @{ Name = "PM_NotificationQueue";      Category = "Notification" },
    @{ Name = "PM_ReminderSchedule";       Category = "Notification" },

    # Admin & Config
    @{ Name = "PM_Configuration";          Category = "Admin" },
    @{ Name = "PM_UserProfiles";           Category = "Admin" },
    @{ Name = "PM_EmailTemplates";         Category = "Admin" },

    # Policy Packs
    @{ Name = "PM_PolicyPacks";            Category = "Packs" },
    @{ Name = "PM_PolicyPackAssignments";  Category = "Packs" },

    # Analytics & Audit
    @{ Name = "PM_PolicyAuditLog";         Category = "Audit" },
    @{ Name = "PM_PolicyAnalytics";        Category = "Audit" },
    @{ Name = "PM_PolicyFeedback";         Category = "Audit" }
)

# Optional / V2 lists (warn if missing, don't fail)
$optionalLists = @(
    @{ Name = "PM_PolicyRatings";          Category = "Social (V2)" },
    @{ Name = "PM_PolicyComments";         Category = "Social (V2)" },
    @{ Name = "PM_PolicyCommentLikes";     Category = "Social (V2)" },
    @{ Name = "PM_PolicyShares";           Category = "Social (V2)" },
    @{ Name = "PM_PolicyFollowers";        Category = "Social (V2)" },
    @{ Name = "PM_WorkflowTemplates";      Category = "Workflow (V2)" },
    @{ Name = "PM_WorkflowInstances";      Category = "Workflow (V2)" },
    @{ Name = "PM_ApprovalChains";         Category = "Workflow (V2)" },
    @{ Name = "PM_ApprovalTemplates";      Category = "Workflow (V2)" },
    @{ Name = "PM_RetentionPolicies";      Category = "Retention (V2)" },
    @{ Name = "PM_LegalHolds";            Category = "Retention (V2)" },
    @{ Name = "PM_SLABreaches";           Category = "Retention (V2)" }
)

foreach ($listDef in $requiredLists) {
    $list = Get-PnPList -Identity $listDef.Name -ErrorAction SilentlyContinue
    if ($list) {
        $itemCount = $list.ItemCount
        Test-Check "Lists" "$($listDef.Name)" $true "($itemCount items)"
    } else {
        Test-Check "Lists" "$($listDef.Name)" $false "List not found"
    }
}

Write-Host ""
Write-Host "  Optional / V2 Lists:" -ForegroundColor Gray

foreach ($listDef in $optionalLists) {
    $list = Get-PnPList -Identity $listDef.Name -ErrorAction SilentlyContinue
    if ($list) {
        Write-Host "  [OK]   $($listDef.Name) ($($list.ItemCount) items)" -ForegroundColor Gray
    } else {
        Test-Warning "$($listDef.Name)" "Not provisioned (optional - $($listDef.Category))"
    }
}

# ============================================================================
# CHECK 2: PM_Configuration Required Keys
# ============================================================================

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  CHECK 2: PM_Configuration Required Keys" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$requiredKeys = @(
    "Integration.AI.Chat.Enabled",
    "Admin.Branding.CompanyName",
    "Admin.Branding.ProductName",
    "Admin.Upload.DocLimitMB",
    "Admin.Quiz.DefaultPassingScore",
    "Compliance.RequireAcknowledgement",
    "Compliance.DefaultDeadlineDays",
    "Compliance.ReviewFrequencyMonths",
    "Approval.RequireOnNew",
    "Approval.AllowSelfApproval",
    "Notifications.NewPolicy.Enabled"
)

$configList = Get-PnPList -Identity "PM_Configuration" -ErrorAction SilentlyContinue
if ($configList) {
    $configItems = Get-PnPListItem -List "PM_Configuration" -PageSize 500 -ErrorAction SilentlyContinue
    $configKeys = @()
    if ($configItems) {
        $configKeys = $configItems | ForEach-Object { $_.FieldValues["ConfigKey"] } | Where-Object { $_ }
    }

    foreach ($key in $requiredKeys) {
        $found = $configKeys -contains $key
        Test-Check "Config" "ConfigKey: $key" $found $(if (-not $found) { "Key not found in PM_Configuration" } else { "" })
    }

    # Check for Azure function URLs (warn if missing)
    $azureKeys = @(
        "Integration.AI.Chat.FunctionUrl",
        "Integration.AI.Quiz.FunctionUrl",
        "Integration.DocConverter.FunctionUrl"
    )
    foreach ($key in $azureKeys) {
        $found = $configKeys -contains $key
        if (-not $found) {
            Test-Warning "ConfigKey: $key" "Not configured (Azure function URL - set after Azure deployment)"
        } else {
            Write-Host "  [OK]   ConfigKey: $key" -ForegroundColor Green
        }
    }
} else {
    Test-Check "Config" "PM_Configuration list exists" $false "List not found"
}

# ============================================================================
# CHECK 3: PM_UserProfiles Population
# ============================================================================

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  CHECK 3: PM_UserProfiles Population" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$userList = Get-PnPList -Identity "PM_UserProfiles" -ErrorAction SilentlyContinue
if ($userList) {
    $userCount = $userList.ItemCount
    Test-Check "Users" "PM_UserProfiles exists" $true "$userCount user(s)"

    if ($userCount -eq 0) {
        Test-Warning "PM_UserProfiles is empty" "Import users from CSV or sync from Entra ID before go-live"
    } else {
        # Check for at least one Admin user
        $adminUsers = Get-PnPListItem -List "PM_UserProfiles" -Query "<View><Query><Where><Eq><FieldRef Name='PMRole'/><Value Type='Text'>Admin</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
        $adminCount = if ($adminUsers) { @($adminUsers).Count } else { 0 }
        Test-Check "Users" "At least 1 Admin user exists" ($adminCount -gt 0) $(if ($adminCount -eq 0) { "No Admin users found" } else { "$adminCount Admin user(s)" })
    }
} else {
    Test-Check "Users" "PM_UserProfiles exists" $false "List not found"
}

# ============================================================================
# CHECK 4: PM_Policies (if seed data was run)
# ============================================================================

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  CHECK 4: PM_Policies Data" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$policyList = Get-PnPList -Identity "PM_Policies" -ErrorAction SilentlyContinue
if ($policyList) {
    $policyCount = $policyList.ItemCount
    Test-Check "Policies" "PM_Policies exists" $true "$policyCount policies"

    if ($policyCount -eq 0) {
        Test-Warning "PM_Policies is empty" "Run sample data scripts or create policies manually"
    }
} else {
    Test-Check "Policies" "PM_Policies exists" $false "List not found"
}

# ============================================================================
# CHECK 5: SharePoint Pages
# ============================================================================

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  CHECK 5: SharePoint Pages" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$requiredPages = @(
    "PolicyHub",
    "MyPolicies",
    "PolicyAdmin",
    "PolicyBuilder",
    "PolicyAuthor",
    "PolicyDetails",
    "PolicyPacks",
    "QuizBuilder",
    "PolicySearch",
    "PolicyHelp",
    "PolicyDistribution",
    "PolicyAnalytics",
    "PolicyManagerView",
    "PolicyAuthorReports",
    "PolicyBulkUpload"
)

foreach ($pageName in $requiredPages) {
    $page = $null
    try {
        $page = Get-PnPPage -Identity $pageName -ErrorAction SilentlyContinue
    } catch {
        # Page doesn't exist
    }

    if ($page) {
        Test-Check "Pages" "$pageName.aspx" $true
    } else {
        Test-Check "Pages" "$pageName.aspx" $false "Page not found"
    }
}

# ============================================================================
# CHECK 6: Key List Fields (spot check critical columns)
# ============================================================================

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "  CHECK 6: Critical Field Spot Checks" -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Cyan

$fieldChecks = @(
    @{ List = "PM_Policies"; Field = "PolicyNumber" },
    @{ List = "PM_Policies"; Field = "PolicyName" },
    @{ List = "PM_Policies"; Field = "PolicyStatus" },
    @{ List = "PM_Policies"; Field = "PolicyCategory" },
    @{ List = "PM_Policies"; Field = "VersionNumber" },
    @{ List = "PM_Policies"; Field = "HTMLContent" },
    @{ List = "PM_Policies"; Field = "Visibility" },
    @{ List = "PM_Policies"; Field = "SubCategory" },
    @{ List = "PM_PolicyAcknowledgements"; Field = "AckUserId" },
    @{ List = "PM_PolicyAcknowledgements"; Field = "AckStatus" },
    @{ List = "PM_Configuration"; Field = "ConfigKey" },
    @{ List = "PM_Configuration"; Field = "ConfigValue" },
    @{ List = "PM_Configuration"; Field = "Category" },
    @{ List = "PM_UserProfiles"; Field = "Email" },
    @{ List = "PM_UserProfiles"; Field = "PMRole" },
    @{ List = "PM_UserProfiles"; Field = "Department" },
    @{ List = "PM_NotificationQueue"; Field = "QueueStatus" }
)

foreach ($check in $fieldChecks) {
    $list = Get-PnPList -Identity $check.List -ErrorAction SilentlyContinue
    if ($list) {
        $field = Get-PnPField -List $check.List -Identity $check.Field -ErrorAction SilentlyContinue
        if ($field) {
            Write-Host "  [OK]   $($check.List).$($check.Field)" -ForegroundColor Gray
        } else {
            Test-Check "Fields" "$($check.List).$($check.Field)" $false "Field not found"
        }
    } else {
        Write-Host "  [SKIP] $($check.List).$($check.Field) (list not found)" -ForegroundColor Yellow
    }
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  VERIFICATION SUMMARY" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Total Checks: $totalChecks" -ForegroundColor White
Write-Host "  Passed:       $passed" -ForegroundColor Green
Write-Host "  Failed:       $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
Write-Host "  Warnings:     $warnings" -ForegroundColor $(if ($warnings -gt 0) { "Yellow" } else { "Green" })
Write-Host ""

$passRate = if ($totalChecks -gt 0) { [math]::Round(($passed / $totalChecks) * 100, 1) } else { 0 }
Write-Host "  Pass Rate:    $passRate%" -ForegroundColor $(if ($passRate -ge 90) { "Green" } elseif ($passRate -ge 70) { "Yellow" } else { "Red" })
Write-Host ""

if ($failedItems.Count -gt 0) {
    Write-Host "  FAILED ITEMS:" -ForegroundColor Red
    foreach ($item in $failedItems) {
        Write-Host "    - $item" -ForegroundColor Red
    }
    Write-Host ""
}

if ($failed -eq 0) {
    Write-Host "  RESULT: ALL CHECKS PASSED" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Deployment verification successful." -ForegroundColor Green
    Write-Host "  Proceed to manual testing (see Deployment Guide Section 9.2)." -ForegroundColor White
} elseif ($failed -le 3) {
    Write-Host "  RESULT: MOSTLY PASSED (minor issues)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Review the failed items above and re-run the relevant" -ForegroundColor Yellow
    Write-Host "  provisioning scripts to resolve." -ForegroundColor Yellow
} else {
    Write-Host "  RESULT: DEPLOYMENT INCOMPLETE" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Multiple checks failed. Re-run the master provisioning" -ForegroundColor Red
    Write-Host "  script (00-Master-Provisioning.ps1) and review errors." -ForegroundColor Red
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
