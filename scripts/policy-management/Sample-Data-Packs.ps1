# ============================================================================
# JML Policy Management - Sample Data: Policy Packs (Simplified)
# Creates onboarding and role-based policy bundles
# ============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"
)

$clientId = "d91b5b78-de72-424e-898b-8b5c9512ebd9"
$tenantId = "03bbbdee-d78b-4613-9b99-c468398246b7"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  JML Policy Management - Policy Packs" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

Import-Module PnP.PowerShell -ErrorAction Stop
Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $clientId -Tenant $tenantId
Write-Host "Connected!" -ForegroundColor Green

# ============================================================================
# POLICY PACKS - Simplified to use only basic fields
# ============================================================================

$policyPacks = @(
    @{
        Title = "New Employee Onboarding - Day 1 Essential Policies"
        PackDescription = "Critical policies that all new employees must read and acknowledge on their first day."
    },
    @{
        Title = "New Employee Onboarding - Week 1 Policies"
        PackDescription = "Additional policies to be completed within your first week."
    },
    @{
        Title = "New Employee Onboarding - Month 1 Policies"
        PackDescription = "Remaining onboarding policies to complete within your first month."
    },
    @{
        Title = "IT Department Policies"
        PackDescription = "Comprehensive policy pack for IT department staff."
    },
    @{
        Title = "Finance Team Policies"
        PackDescription = "Essential policies for finance department staff."
    },
    @{
        Title = "Manager Essential Policies"
        PackDescription = "Additional policies for people managers."
    },
    @{
        Title = "Internal Transfer - Role Change Policies"
        PackDescription = "Policies to acknowledge when changing roles internally."
    },
    @{
        Title = "Leaver Acknowledgement Pack"
        PackDescription = "Policies that departing employees must acknowledge."
    },
    @{
        Title = "Annual Compliance Refresh"
        PackDescription = "Annual re-acknowledgement of critical compliance policies."
    }
)

Write-Host "`n[1/1] Creating policy packs..." -ForegroundColor Yellow

foreach ($pack in $policyPacks) {
    try {
        Add-PnPListItem -List "PM_PolicyPacks" -Values $pack | Out-Null
        Write-Host "  Created: $($pack.Title)" -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed: $($pack.Title) - $_" -ForegroundColor Red
    }
}

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Policy Packs created: $($policyPacks.Count)" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan

Disconnect-PnPOnline
