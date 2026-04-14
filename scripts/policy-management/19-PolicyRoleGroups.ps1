# ============================================================================
# Policy Manager - Role-Based SharePoint Groups
# Creates 3 PM_ security groups used by RoleDetectionService for role mapping.
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Policy Manager Role Groups Provisioning" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$groups = @(
    @{
        Name        = "PM_PolicyAdmins"
        Description = "Policy Manager Administrators — full system access, all configuration and user management"
    },
    @{
        Name        = "PM_PolicyManagers"
        Description = "Policy Manager Managers — analytics, approvals, distribution, SLA oversight, team compliance"
    },
    @{
        Name        = "PM_PolicyAuthors"
        Description = "Policy Manager Authors — create and edit policies, manage packs, quiz builder access"
    }
)

$index = 0
foreach ($grp in $groups) {
    $index++
    Write-Host "[$index/$($groups.Count)] Creating $($grp.Name)..." -ForegroundColor Yellow

    $existing = Get-PnPGroup -Identity $grp.Name -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        New-PnPGroup -Title $grp.Name -Description $grp.Description | Out-Null
        Write-Host "  Created: $($grp.Name)" -ForegroundColor Green
    } else {
        Write-Host "  Exists: $($grp.Name)" -ForegroundColor Gray
    }
}

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  POLICY ROLE GROUPS - COMPLETE" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  3 groups created/verified:" -ForegroundColor White
Write-Host "    PM_PolicyAdmins    — Maps to Admin role (full access)" -ForegroundColor Gray
Write-Host "    PM_PolicyManagers  — Maps to Manager role (oversight)" -ForegroundColor Gray
Write-Host "    PM_PolicyAuthors   — Maps to Author role (create/edit)" -ForegroundColor Gray
Write-Host ""
Write-Host "  Role detection (single source of truth):" -ForegroundColor White
Write-Host "    1. PM_UserProfiles.PMRole — set in Admin Centre > User Directory" -ForegroundColor Gray
Write-Host "    2. Site Collection Admin — gets Admin if no PMRole set" -ForegroundColor Gray
Write-Host "    3. Default — User role (no entry = basic access)" -ForegroundColor Gray
Write-Host ""
Write-Host "  SP groups are synced FOR REFERENCE when admin assigns roles." -ForegroundColor White
Write-Host "  They are NOT used for role detection — PM_UserProfiles is the source." -ForegroundColor White
Write-Host ""
Write-Host "  NOTE: Everyone without a PM_UserProfiles entry = User role." -ForegroundColor Yellow
Write-Host "  Site Collection Admins get Admin UNLESS explicitly given" -ForegroundColor Yellow
Write-Host "  a lower role in Admin Centre (e.g. Author)." -ForegroundColor Yellow
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
