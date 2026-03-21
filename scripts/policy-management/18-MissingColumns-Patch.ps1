# ============================================================================
# Policy Manager — Missing Columns Patch
# Adds columns that the code references but don't exist on the SP lists
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Missing Columns Patch" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# PM_PolicyPacks — TargetProcessType column
$listName = "PM_PolicyPacks"
Write-Host "Adding TargetProcessType to $listName..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "TargetProcessType" -InternalName "TargetProcessType" -Type Text -ErrorAction SilentlyContinue | Out-Null
Write-Host "  Done" -ForegroundColor Gray

# PM_Policies — missing audience/review columns from Policy Builder fixes
$listName = "PM_Policies"
Write-Host "Adding audience + review columns to $listName..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "Departments" -InternalName "Departments" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TargetRoles" -InternalName "TargetRoles" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TargetLocations" -InternalName "TargetLocations" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IncludeContractors" -InternalName "IncludeContractors" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ReviewFrequency" -InternalName "ReviewFrequency" -Type Choice -Choices "Annual","Biannual","Quarterly","Monthly","None" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "SupersedesPolicy" -InternalName "SupersedesPolicy" -Type Text -ErrorAction SilentlyContinue | Out-Null
Write-Host "  Done" -ForegroundColor Gray

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Patch complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
