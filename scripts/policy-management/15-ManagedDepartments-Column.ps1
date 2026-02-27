# ============================================================================
# Policy Manager - Add ManagedDepartments Column to PM_Employees
# Adds a multi-value text column for assigning users to multiple departments.
# Use case: A single manager responsible for HR, Finance, and Legal.
# Values stored as semicolon-delimited string (e.g., "HR;Finance;Legal")
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Adding ManagedDepartments Column to PM_Employees" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$listName = "PM_Employees"

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    Write-Host "  [ERROR] $listName does not exist. Run 12-UserManagement-Lists.ps1 first." -ForegroundColor Red
    return
}

Write-Host "  Adding ManagedDepartments column..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "ManagedDepartments" -InternalName "ManagedDepartments" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "  Done: ManagedDepartments (Note/multi-line text)" -ForegroundColor Green

Write-Host ""
Write-Host "  Column added to $listName." -ForegroundColor Green
Write-Host "  Values are semicolon-delimited (e.g., 'HR;Finance;Legal')." -ForegroundColor Gray
Write-Host ""
