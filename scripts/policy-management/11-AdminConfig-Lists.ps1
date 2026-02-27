# ============================================================================
# Policy Manager - Admin Configuration Lists
# Creates 5 admin config lists: NamingRules, SLAConfigs,
# DataLifecyclePolicies, EmailTemplates, PolicyCategories
# Also patches PM_PolicyMetadataProfiles with missing columns
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Admin Configuration Lists Provisioning" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ----------------------------------------------------------------------------
# 1. PM_NamingRules
# ----------------------------------------------------------------------------
$listName = "PM_NamingRules"
Write-Host "[1/6] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Pattern" -InternalName "Pattern" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Segments" -InternalName "Segments" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AppliesTo" -InternalName "AppliesTo" -Type Choice -Choices "All Policies","HR Policies","Compliance Policies","IT Policies","Finance Policies","Legal Policies","Operational Policies" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Example" -InternalName "Example" -Type Text -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# ----------------------------------------------------------------------------
# 2. PM_SLAConfigs
# ----------------------------------------------------------------------------
$listName = "PM_SLAConfigs"
Write-Host "[2/6] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "ProcessType" -InternalName "ProcessType" -Type Choice -Choices "Review","Acknowledgement","Approval","Authoring","Audit","Distribution","Escalation" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TargetDays" -InternalName "TargetDays" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "WarningThresholdDays" -InternalName "WarningThresholdDays" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# ----------------------------------------------------------------------------
# 3. PM_DataLifecyclePolicies
# ----------------------------------------------------------------------------
$listName = "PM_DataLifecyclePolicies"
Write-Host "[3/6] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "EntityType" -InternalName "EntityType" -Type Choice -Choices "Policies","Drafts","Acknowledgements","AuditLogs","Approvals","Quizzes","Notifications","Analytics" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RetentionPeriodDays" -InternalName "RetentionPeriodDays" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AutoDeleteEnabled" -InternalName "AutoDeleteEnabled" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ArchiveBeforeDelete" -InternalName "ArchiveBeforeDelete" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# ----------------------------------------------------------------------------
# 4. PM_EmailTemplates
# ----------------------------------------------------------------------------
$listName = "PM_EmailTemplates"
Write-Host "[4/6] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "EventTrigger" -InternalName "EventTrigger" -Type Choice -Choices "Policy Published","Ack Overdue","Approval Needed","Policy Expiring","SLA Breached","Violation Found","Campaign Active","User Added","Policy Updated","Policy Retired" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Subject" -InternalName "Subject" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Body" -InternalName "Body" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Recipients" -InternalName "Recipients" -Type Choice -Choices "All Employees","Assigned Users","Approvers","Policy Owners","Managers","Compliance Officers","Target Groups","New Users","HR Team","IT Admins" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "MergeTags" -InternalName "MergeTags" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# ----------------------------------------------------------------------------
# 5. PM_PolicyCategories
# ----------------------------------------------------------------------------
$listName = "PM_PolicyCategories"
Write-Host "[5/6] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "CategoryName" -InternalName "CategoryName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IconName" -InternalName "IconName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Color" -InternalName "Color" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "SortOrder" -InternalName "SortOrder" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsDefault" -InternalName "IsDefault" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# Seed default categories
$existingItems = Get-PnPListItem -List $listName -ErrorAction SilentlyContinue
if ($null -eq $existingItems -or $existingItems.Count -eq 0) {
    Write-Host "    Seeding default categories..." -ForegroundColor Yellow
    $defaults = @(
        @{ Title = "HR Policies"; CategoryName = "HR Policies"; IconName = "People"; Color = "#0d9488"; SortOrder = 1; IsActive = $true; IsDefault = $true; Description = "Human resources policies" },
        @{ Title = "IT & Security"; CategoryName = "IT & Security"; IconName = "Shield"; Color = "#2563eb"; SortOrder = 2; IsActive = $true; IsDefault = $true; Description = "IT security and technology policies" },
        @{ Title = "Health & Safety"; CategoryName = "Health & Safety"; IconName = "Health"; Color = "#059669"; SortOrder = 3; IsActive = $true; IsDefault = $true; Description = "Workplace health and safety policies" },
        @{ Title = "Compliance"; CategoryName = "Compliance"; IconName = "ComplianceAudit"; Color = "#7c3aed"; SortOrder = 4; IsActive = $true; IsDefault = $true; Description = "Regulatory compliance policies" },
        @{ Title = "Financial"; CategoryName = "Financial"; IconName = "Money"; Color = "#d97706"; SortOrder = 5; IsActive = $true; IsDefault = $true; Description = "Financial and accounting policies" },
        @{ Title = "Operational"; CategoryName = "Operational"; IconName = "Settings"; Color = "#475569"; SortOrder = 6; IsActive = $true; IsDefault = $true; Description = "Operational procedures and policies" },
        @{ Title = "Legal"; CategoryName = "Legal"; IconName = "Library"; Color = "#be185d"; SortOrder = 7; IsActive = $true; IsDefault = $true; Description = "Legal and contractual policies" },
        @{ Title = "Environmental"; CategoryName = "Environmental"; IconName = "Leaf"; Color = "#16a34a"; SortOrder = 8; IsActive = $true; IsDefault = $true; Description = "Environmental sustainability policies" },
        @{ Title = "Quality Assurance"; CategoryName = "Quality Assurance"; IconName = "CheckboxComposite"; Color = "#0891b2"; SortOrder = 9; IsActive = $true; IsDefault = $true; Description = "Quality assurance and standards" },
        @{ Title = "Data Privacy"; CategoryName = "Data Privacy"; IconName = "LockSolid"; Color = "#dc2626"; SortOrder = 10; IsActive = $true; IsDefault = $true; Description = "Data privacy and protection policies" }
    )
    foreach ($cat in $defaults) {
        Add-PnPListItem -List $listName -Values $cat | Out-Null
    }
    Write-Host "    Seeded ${($defaults.Count)} default categories" -ForegroundColor Green
} else {
    Write-Host "    Categories already seeded ($($existingItems.Count) items)" -ForegroundColor Gray
}

# ----------------------------------------------------------------------------
# 6. PM_PolicyMetadataProfiles — ensure required columns exist
# ----------------------------------------------------------------------------
$listName = "PM_PolicyMetadataProfiles"
Write-Host "[6/6] Patching $listName columns..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "ProfileName" -InternalName "ProfileName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PolicyCategory" -InternalName "PolicyCategory" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ComplianceRisk" -InternalName "ComplianceRisk" -Type Choice -Choices "Critical","High","Medium","Low","Informational" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ReadTimeframe" -InternalName "ReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6","Custom" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RequiresAcknowledgement" -InternalName "RequiresAcknowledgement" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RequiresQuiz" -InternalName "RequiresQuiz" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TargetDepartments" -InternalName "TargetDepartments" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TargetRoles" -InternalName "TargetRoles" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields ensured on $listName" -ForegroundColor Gray

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  ADMIN CONFIG LISTS - COMPLETE" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  6 lists created/verified:" -ForegroundColor White
Write-Host "    PM_NamingRules            — Naming convention rules" -ForegroundColor Gray
Write-Host "    PM_SLAConfigs             — SLA target configurations" -ForegroundColor Gray
Write-Host "    PM_DataLifecyclePolicies  — Data retention policies" -ForegroundColor Gray
Write-Host "    PM_EmailTemplates         — Email notification templates" -ForegroundColor Gray
Write-Host "    PM_PolicyCategories       — Policy category definitions" -ForegroundColor Gray
Write-Host "    PM_PolicyMetadataProfiles — Patched with missing columns" -ForegroundColor Gray
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
