# ============================================================================
# Policy Manager - SubCategory, Folders, Quiz & Request Fields
# Creates PM_PolicySubCategories list and adds fields to PM_Policies
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SubCategory, Quiz & Request Fields Provisioning" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ----------------------------------------------------------------------------
# 1. Add SubCategory column to PM_Policies
# ----------------------------------------------------------------------------
$listName = "PM_Policies"
Write-Host "[1/4] Adding SubCategory column to $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    Write-Host "  ERROR: $listName does not exist. Run 01-Core-PolicyLists.ps1 first." -ForegroundColor Red
} else {
    Add-PnPField -List $listName -DisplayName "SubCategory" -InternalName "SubCategory" -Type Text -ErrorAction SilentlyContinue | Out-Null
    Add-PnPField -List $listName -DisplayName "LinkedQuizId" -InternalName "LinkedQuizId" -Type Number -ErrorAction SilentlyContinue | Out-Null
    Add-PnPField -List $listName -DisplayName "SourceRequestId" -InternalName "SourceRequestId" -Type Number -ErrorAction SilentlyContinue | Out-Null
    Write-Host "  Done: SubCategory (Text), LinkedQuizId (Number), SourceRequestId (Number)" -ForegroundColor Gray
}

# ----------------------------------------------------------------------------
# 2. Create PM_PolicySubCategories list
# ----------------------------------------------------------------------------
$listName = "PM_PolicySubCategories"
Write-Host "[2/4] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "SubCategoryName" -InternalName "SubCategoryName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ParentCategoryId" -InternalName "ParentCategoryId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ParentCategoryName" -InternalName "ParentCategoryName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IconName" -InternalName "IconName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "SortOrder" -InternalName "SortOrder" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Write-Host "  Fields added to $listName" -ForegroundColor Gray

# ----------------------------------------------------------------------------
# 3. Ensure PM_PolicySourceDocuments library exists
# ----------------------------------------------------------------------------
$docLibName = "PM_PolicySourceDocuments"
Write-Host "[3/4] Verifying $docLibName library..." -ForegroundColor Yellow

$docLib = Get-PnPList -Identity $docLibName -ErrorAction SilentlyContinue
if ($null -eq $docLib) {
    Write-Host "  WARNING: $docLibName does not exist. Run 09-PolicySourceDocuments.ps1 first." -ForegroundColor Yellow
} else {
    Write-Host "  Exists: $docLibName" -ForegroundColor Gray
}

# ----------------------------------------------------------------------------
# 4. Summary
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SUBCATEGORY & FOLDERS - COMPLETE" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Changes:" -ForegroundColor White
Write-Host "    PM_Policies         — Added SubCategory, LinkedQuizId, SourceRequestId columns" -ForegroundColor Gray
Write-Host "    PM_PolicySubCategories — New list with 7 columns" -ForegroundColor Gray
Write-Host ""
Write-Host "  Per-policy document folders will be created automatically" -ForegroundColor Yellow
Write-Host "  in PM_PolicySourceDocuments when policies are saved." -ForegroundColor Yellow
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
