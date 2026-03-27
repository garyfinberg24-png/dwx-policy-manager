# ============================================================================
# Policy Manager — Distribution Missing Columns Patch
# Adds columns that PolicyDistribution.tsx writes but aren't in the original
# provisioning script. Idempotent — safe to re-run.
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  PM_PolicyDistributions — Missing Columns Patch" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$listName = "PM_PolicyDistributions"

# Check list exists
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    Write-Host "  ERROR: List $listName does not exist. Run 03-Exemption-Distribution-Lists.ps1 first." -ForegroundColor Red
    return
}

# CampaignName — the code writes this alongside DistributionName and Title
Write-Host "  Adding CampaignName..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "Campaign Name" -InternalName "CampaignName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Done" -ForegroundColor Gray

# ContentType — Policy or PolicyPack (stored as text, not SP content type)
Write-Host "  Adding ContentType..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "Content Type" -InternalName "ContentType" -Type Choice -Choices "Policy","PolicyPack" -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Done" -ForegroundColor Gray

# TargetUsers — comma-separated user emails/names
Write-Host "  Adding TargetUsers..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "Target Users" -InternalName "TargetUsers" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Done" -ForegroundColor Gray

# TargetGroups — comma-separated group names
Write-Host "  Adding TargetGroups..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "Target Groups" -InternalName "TargetGroups" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Done" -ForegroundColor Gray

# PolicyTitle — display name of the linked policy
Write-Host "  Adding PolicyTitle..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "Policy Title" -InternalName "PolicyTitle" -Type Text -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Done" -ForegroundColor Gray

# PolicyPackId — ID of linked policy pack
Write-Host "  Adding PolicyPackId..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "Policy Pack ID" -InternalName "PolicyPackId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Done" -ForegroundColor Gray

# PolicyPackName — display name of linked policy pack
Write-Host "  Adding PolicyPackName..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "Policy Pack Name" -InternalName "PolicyPackName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Done" -ForegroundColor Gray

# DistributionStatus — tracks campaign lifecycle
Write-Host "  Adding DistributionStatus..." -ForegroundColor Yellow
Add-PnPField -List $listName -DisplayName "Distribution Status" -InternalName "DistributionStatus" -Type Choice -Choices "Draft","Scheduled","Active","Paused","Completed","Cancelled" -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "DistributionStatus" -Values @{DefaultValue="Draft"} -ErrorAction SilentlyContinue
Write-Host "    Done" -ForegroundColor Gray

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Patch complete — 8 columns added to $listName" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
