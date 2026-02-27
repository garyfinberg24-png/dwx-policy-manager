# ============================================================================
# Policy Manager - Visibility & Security Columns
# Adds Visibility and TargetSecurityGroups columns to PM_Policies
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Visibility & Security Columns Provisioning" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$listName = "PM_Policies"

# Check list exists
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    Write-Host "  ERROR: $listName does not exist. Run 01-Core-PolicyLists.ps1 first." -ForegroundColor Red
    return
}

Write-Host "[1/2] Adding Visibility column..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "Visibility" -InternalName "Visibility" -Type Choice -Choices "All Employees","Department","Role","Security Group","Custom" -ErrorAction SilentlyContinue | Out-Null
# Set default value via Set-PnPField (Add-PnPField does not support -DefaultValue)
Set-PnPField -List $listName -Identity "Visibility" -Values @{DefaultValue="All Employees"} -ErrorAction SilentlyContinue
Write-Host "  Done: Visibility (Choice, default: All Employees)" -ForegroundColor Gray

Write-Host "[2/2] Adding TargetSecurityGroups column..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "TargetSecurityGroups" -InternalName "TargetSecurityGroups" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "  Done: TargetSecurityGroups (Note — JSON array of group names)" -ForegroundColor Gray

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  VISIBILITY COLUMNS - COMPLETE" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  2 columns added to PM_Policies:" -ForegroundColor White
Write-Host "    Visibility           — Choice (All Employees, Department, Role, Security Group, Custom)" -ForegroundColor Gray
Write-Host "    TargetSecurityGroups — Note (JSON array of security group names)" -ForegroundColor Gray
Write-Host ""
Write-Host "  NOTE: Existing policies default to 'All Employees' (visible to everyone)." -ForegroundColor Yellow
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
