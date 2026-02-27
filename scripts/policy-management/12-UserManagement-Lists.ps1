# ============================================================================
# Policy Manager - User Management Lists
# Creates 3 user management lists: PM_Employees, PM_Sync_Log, PM_Audiences
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  User Management Lists Provisioning" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ----------------------------------------------------------------------------
# 1. PM_Employees
# ----------------------------------------------------------------------------
$listName = "PM_Employees"
Write-Host "[1/3] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "FirstName" -InternalName "FirstName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LastName" -InternalName "LastName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Email" -InternalName "Email" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "EmployeeNumber" -InternalName "EmployeeNumber" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "JobTitle" -InternalName "JobTitle" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Department" -InternalName "Department" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Location" -InternalName "Location" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "OfficePhone" -InternalName "OfficePhone" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "MobilePhone" -InternalName "MobilePhone" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ManagerEmail" -InternalName "ManagerEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Active","Inactive","PreHire","OnLeave","Terminated","Retired" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "EmploymentType" -InternalName "EmploymentType" -Type Choice -Choices "Full-Time","Part-Time","Contractor","Intern","Temporary" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "CostCenter" -InternalName "CostCenter" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "EntraObjectId" -InternalName "EntraObjectId" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PMRole" -InternalName "PMRole" -Type Choice -Choices "User","Author","Manager","Admin" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ProfilePhoto" -InternalName "ProfilePhoto" -Type URL -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LastSyncedAt" -InternalName "LastSyncedAt" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Notes" -InternalName "Notes" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# ----------------------------------------------------------------------------
# 2. PM_Sync_Log
# ----------------------------------------------------------------------------
$listName = "PM_Sync_Log"
Write-Host "[2/3] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "SyncId" -InternalName "SyncId" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Started","Running","Completed","CompletedWithErrors","Failed" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Message" -InternalName "Message" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# ----------------------------------------------------------------------------
# 3. PM_Audiences
# ----------------------------------------------------------------------------
$listName = "PM_Audiences"
Write-Host "[3/3] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Criteria" -InternalName "Criteria" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "MemberCount" -InternalName "MemberCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LastEvaluated" -InternalName "LastEvaluated" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  USER MANAGEMENT LISTS - COMPLETE" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  3 lists created/verified:" -ForegroundColor White
Write-Host "    PM_Employees    — User directory (synced from Entra ID)" -ForegroundColor Gray
Write-Host "    PM_Sync_Log     — Entra ID sync operation logs" -ForegroundColor Gray
Write-Host "    PM_Audiences    — Custom audience definitions for targeting" -ForegroundColor Gray
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
