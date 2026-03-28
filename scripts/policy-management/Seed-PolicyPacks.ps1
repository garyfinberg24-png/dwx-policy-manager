# =============================================================================
# Seed-PolicyPacks.ps1
# Seeds 8 realistic business-relevant policy packs into PM_PolicyPacks
# Assumes: Already connected to SharePoint via Connect-PnPOnline
# Site: https://mf7m.sharepoint.com/sites/PolicyManager
# =============================================================================

$listName = "PM_PolicyPacks"

# Get existing published policies for realistic pack composition
Write-Host "Loading published policies from PM_Policies..." -ForegroundColor Cyan
$policies = Get-PnPListItem -List "PM_Policies" -Fields "Id","Title","PolicyName","PolicyNumber","PolicyCategory","ComplianceRisk" | Where-Object { $_.FieldValues["PolicyStatus"] -eq "Published" -or $_.FieldValues["PolicyStatus"] -eq "Approved" -or $_.FieldValues["PolicyStatus"] -eq "Draft" }

$policyIds = $policies | ForEach-Object { $_.Id }
Write-Host "Found $($policyIds.Count) policies to distribute across packs" -ForegroundColor Green

# If no policies exist, use placeholder IDs
if ($policyIds.Count -eq 0) {
    Write-Host "No policies found — using placeholder IDs 1-10" -ForegroundColor Yellow
    $policyIds = 1..10
}

# Helper: pick N random policy IDs
function Get-RandomPolicyIds {
    param([int]$Count)
    $available = $policyIds | Get-Random -Count ([Math]::Min($Count, $policyIds.Count))
    return $available
}

# =============================================================================
# 8 Business-Relevant Policy Packs
# =============================================================================

$packs = @(
    @{
        PackName        = "New Employee Onboarding Pack"
        PackDescription = "Essential policies every new employee must read and acknowledge within their first week. Covers workplace conduct, IT security, data protection, and health & safety fundamentals."
        PackType        = "Onboarding"
        PackCategory    = "Onboarding"
        PolicyCount     = 5
        IsSequential    = $true
        IsMandatory     = $true
        SendWelcomeEmail = $true
        SendTeamsNotification = $true
    },
    @{
        PackName        = "IT & Cybersecurity Essentials"
        PackDescription = "Core IT security policies including acceptable use, password management, data classification, remote access, and incident reporting. Required for all staff with system access."
        PackType        = "Department"
        PackCategory    = "Department"
        PolicyCount     = 4
        IsSequential    = $false
        IsMandatory     = $true
        SendWelcomeEmail = $true
        SendTeamsNotification = $true
    },
    @{
        PackName        = "Data Protection & Privacy (POPIA/GDPR)"
        PackDescription = "Comprehensive privacy compliance pack covering data handling, consent management, breach notification, cross-border transfers, and data subject rights."
        PackType        = "Custom"
        PackCategory    = "Custom"
        PolicyCount     = 4
        IsSequential    = $false
        IsMandatory     = $true
        SendWelcomeEmail = $true
        SendTeamsNotification = $false
    },
    @{
        PackName        = "Health, Safety & Wellbeing"
        PackDescription = "Workplace health and safety policies including emergency procedures, ergonomics, mental health support, first aid protocols, and incident reporting."
        PackType        = "Department"
        PackCategory    = "Department"
        PolicyCount     = 3
        IsSequential    = $false
        IsMandatory     = $true
        SendWelcomeEmail = $true
        SendTeamsNotification = $true
    },
    @{
        PackName        = "Finance & Procurement Compliance"
        PackDescription = "Financial governance policies covering expense management, procurement procedures, anti-bribery, conflict of interest, and financial delegation of authority."
        PackType        = "Role"
        PackCategory    = "Role"
        PolicyCount     = 4
        IsSequential    = $false
        IsMandatory     = $true
        SendWelcomeEmail = $false
        SendTeamsNotification = $true
    },
    @{
        PackName        = "Manager Responsibilities Pack"
        PackDescription = "Policies that all people managers must understand: performance management, leave & attendance, grievance handling, disciplinary procedures, and team compliance obligations."
        PackType        = "Role"
        PackCategory    = "Role"
        PolicyCount     = 5
        IsSequential    = $true
        IsMandatory     = $true
        SendWelcomeEmail = $true
        SendTeamsNotification = $true
    },
    @{
        PackName        = "Annual Compliance Refresh"
        PackDescription = "Yearly re-acknowledgement pack for critical compliance policies. Distributed to all employees at the start of each financial year to maintain regulatory compliance."
        PackType        = "Custom"
        PackCategory    = "Custom"
        PolicyCount     = 6
        IsSequential    = $false
        IsMandatory     = $true
        SendWelcomeEmail = $true
        SendTeamsNotification = $true
    },
    @{
        PackName        = "Remote & Hybrid Working"
        PackDescription = "Policies for remote and hybrid workers covering home office setup, VPN usage, communication standards, equipment care, and work-from-home health & safety."
        PackType        = "Location"
        PackCategory    = "Location"
        PolicyCount     = 3
        IsSequential    = $false
        IsMandatory     = $false
        SendWelcomeEmail = $true
        SendTeamsNotification = $false
    }
)

Write-Host "`nSeeding $($packs.Count) policy packs into $listName..." -ForegroundColor Cyan

foreach ($pack in $packs) {
    $selectedIds = Get-RandomPolicyIds -Count $pack.PolicyCount

    $values = @{
        Title           = $pack.PackName
        PackName        = $pack.PackName
        PackDescription = $pack.PackDescription
        PolicyIds       = ($selectedIds | ConvertTo-Json -Compress)
        PolicyCount     = $selectedIds.Count
        IsActive        = $true
    }

    try {
        $item = Add-PnPListItem -List $listName -Values $values
        Write-Host "  [OK] $($pack.PackName) — $($selectedIds.Count) policies" -ForegroundColor Green

        # Phase 2: Optional fields (may not be provisioned)
        try {
            $optionalValues = @{
                PackType              = $pack.PackType
                PackCategory          = $pack.PackCategory
                IsMandatory           = $pack.IsMandatory
                IsSequential          = $pack.IsSequential
                SendWelcomeEmail      = $pack.SendWelcomeEmail
                SendTeamsNotification = $pack.SendTeamsNotification
            }
            if ($pack.IsSequential) {
                $optionalValues["PolicySequence"] = ($selectedIds | ConvertTo-Json -Compress)
            }
            Set-PnPListItem -List $listName -Identity $item.Id -Values $optionalValues | Out-Null
        }
        catch {
            Write-Host "    [WARN] Optional fields skipped: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "  [FAIL] $($pack.PackName): $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`nDone! $($packs.Count) policy packs seeded." -ForegroundColor Green
Write-Host "View at: https://mf7m.sharepoint.com/sites/PolicyManager/Lists/PM_PolicyPacks" -ForegroundColor Cyan
