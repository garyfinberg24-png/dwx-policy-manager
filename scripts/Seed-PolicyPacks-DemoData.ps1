# ============================================================================
# JML Policy Management - Seed Policy Packs Demo Data
# Populates JML_PolicyPacks with realistic South African enterprise data
# Target: https://mf7m.sharepoint.com/sites/JML (Development)
# ============================================================================
#
# USAGE:
#   .\Seed-PolicyPacks-DemoData.ps1
#
# PREREQUISITES:
#   - PnP.PowerShell module installed
#   - JML_PolicyPacks list provisioned (run Provision-PolicyManager-Lists.ps1 first)
#   - JML_Policies list populated with policies (PolicyIds reference these)
#
# This will open a browser for authentication.
# ============================================================================

$SiteUrl = "https://mf7m.sharepoint.com/sites/JML"
$clientId = "d91b5b78-de72-424e-898b-8b5c9512ebd9"
$tenantId = "03bbbdee-d78b-4613-9b99-c468398246b7"

$ErrorActionPreference = "Continue"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  JML Policy Management - Seed Policy Packs Demo Data" -ForegroundColor Cyan
Write-Host "  Target: $SiteUrl" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Check for PnP PowerShell module
$module = Get-Module -ListAvailable -Name "PnP.PowerShell"
if (-not $module) {
    Write-Host "PnP.PowerShell module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

Import-Module PnP.PowerShell -ErrorAction Stop
Write-Host "PnP.PowerShell module loaded" -ForegroundColor Green

# Connect to SharePoint
Write-Host ""
Write-Host "Connecting to SharePoint using Device Login..." -ForegroundColor Cyan
Write-Host "Follow the instructions to authenticate in your browser." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $clientId -Tenant $tenantId
    Write-Host "Connected successfully!" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to SharePoint: $_" -ForegroundColor Red
    exit 1
}

$web = Get-PnPWeb
Write-Host "Connected to: $($web.Title)" -ForegroundColor Green
Write-Host ""

# ============================================================================
# Fetch existing policies to build realistic PolicyIds references
# ============================================================================
Write-Host "Fetching existing policies from JML_Policies..." -ForegroundColor Yellow
$existingPolicies = @()
try {
    $existingPolicies = Get-PnPListItem -List "JML_Policies" -PageSize 500 | ForEach-Object {
        @{
            Id = $_.Id
            PolicyNumber = $_.FieldValues["PolicyNumber"]
            PolicyName = $_.FieldValues["PolicyName"]
            Category = $_.FieldValues["PolicyCategory"]
        }
    }
    Write-Host "  Found $($existingPolicies.Count) policies" -ForegroundColor Green
} catch {
    Write-Host "  Warning: Could not fetch policies. PolicyIds will use placeholder IDs." -ForegroundColor Yellow
}

# Helper: Get random policy IDs from existing policies
function Get-RandomPolicyIds {
    param([int]$Count = 5)
    if ($existingPolicies.Count -ge $Count) {
        $selected = $existingPolicies | Get-Random -Count $Count
        return ($selected | ForEach-Object { $_.Id }) -join ","
    } elseif ($existingPolicies.Count -gt 0) {
        return ($existingPolicies | ForEach-Object { $_.Id }) -join ","
    } else {
        # Fallback placeholder IDs
        return (1..$Count | ForEach-Object { Get-Random -Minimum 1 -Maximum 50 }) -join ","
    }
}

# ============================================================================
# Policy Pack Definitions — Realistic SA Enterprise Data
# ============================================================================

$listName = "JML_PolicyPacks"

$policyPacks = @(
    @{
        PackName = "New Employee Onboarding Pack"
        PackDescription = "Essential policies every new employee at our South African offices must read and acknowledge within their first week. Covers code of conduct, health & safety, IT acceptable use, POPIA data privacy, and workplace ethics."
        PackType = "Onboarding"
        IsActive = $true
        PolicyCount = 8
        IsSequential = $true
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "IT & Cybersecurity Essentials"
        PackDescription = "Mandatory IT security policies for all staff with access to company systems. Includes password management, phishing awareness, BYOD policy, remote access, and incident reporting procedures aligned with SA POPIA requirements."
        PackType = "Department"
        IsActive = $true
        PolicyCount = 6
        IsSequential = $false
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Health & Safety Compliance Pack"
        PackDescription = "Occupational Health and Safety Act (OHSA) compliance pack for all employees working in South African facilities. Covers workplace safety, emergency procedures, PPE requirements, and incident reporting."
        PackType = "Department"
        IsActive = $true
        PolicyCount = 7
        IsSequential = $false
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "POPIA & Data Privacy Pack"
        PackDescription = "Protection of Personal Information Act (POPIA) compliance training pack. Mandatory for all employees handling personal data. Covers data subject rights, consent management, breach notification, and cross-border data transfers."
        PackType = "Custom"
        IsActive = $true
        PolicyCount = 5
        IsSequential = $true
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Financial Services Regulatory Pack"
        PackDescription = "FSCA and SARB regulatory compliance policies for financial services division. Covers FICA, anti-money laundering (AML), FAIS compliance, and financial crime prevention procedures."
        PackType = "Department"
        IsActive = $true
        PolicyCount = 6
        IsSequential = $false
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Management & Leadership Pack"
        PackDescription = "Policies specific to people managers and team leads. Covers performance management, disciplinary procedures (aligned with LRA), leave management, diversity & inclusion, and employee wellbeing responsibilities."
        PackType = "Role"
        IsActive = $true
        PolicyCount = 7
        IsSequential = $false
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Johannesburg Office Induction"
        PackDescription = "Location-specific induction pack for the Sandton head office. Covers building access, parking, visitor management, emergency evacuation routes, and local facility procedures."
        PackType = "Location"
        IsActive = $true
        PolicyCount = 4
        IsSequential = $true
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Cape Town Office Induction"
        PackDescription = "Location-specific induction pack for the Cape Town Century City office. Covers building access, parking, load shedding procedures, emergency evacuation, and local transport arrangements."
        PackType = "Location"
        IsActive = $true
        PolicyCount = 4
        IsSequential = $true
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Durban Branch Induction"
        PackDescription = "Location-specific induction pack for the Durban Umhlanga Ridge branch. Building orientation, parking facilities, security protocols, and coastal weather emergency procedures."
        PackType = "Location"
        IsActive = $true
        PolicyCount = 3
        IsSequential = $true
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "B-BBEE Compliance Pack"
        PackDescription = "Broad-Based Black Economic Empowerment compliance policies. Covers procurement transformation, skills development, enterprise development, and socio-economic development requirements."
        PackType = "Custom"
        IsActive = $true
        PolicyCount = 5
        IsSequential = $false
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Employee Exit & Offboarding Pack"
        PackDescription = "Policies to be acknowledged during the offboarding process. Covers confidentiality post-employment, intellectual property, return of assets, restraint of trade, and exit interview requirements."
        PackType = "Custom"
        IsActive = $true
        PolicyCount = 5
        IsSequential = $true
        TargetProcessType = "Leaver"
    },
    @{
        PackName = "Internal Transfer Pack"
        PackDescription = "Policies for employees moving between departments or roles. Covers role-specific compliance requirements, updated reporting lines, new system access policies, and handover procedures."
        PackType = "Custom"
        IsActive = $true
        PolicyCount = 4
        IsSequential = $false
        TargetProcessType = "Mover"
    },
    @{
        PackName = "Contractor & Temp Staff Pack"
        PackDescription = "Abbreviated policy pack for contractors, temporary staff, and consultants. Covers NDA requirements, site access, IT acceptable use, health & safety basics, and POPIA obligations."
        PackType = "Onboarding"
        IsActive = $true
        PolicyCount = 5
        IsSequential = $true
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Annual Compliance Refresh"
        PackDescription = "Yearly compliance refresh pack assigned to all employees during Q1. Covers updated code of conduct, anti-bribery & corruption, POPIA refresher, workplace harassment, and whistleblowing policy."
        PackType = "Custom"
        IsActive = $true
        PolicyCount = 6
        IsSequential = $false
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Mining Operations Safety Pack"
        PackDescription = "Mine Health and Safety Act (MHSA) compliance pack for employees deployed to mining operation sites. Covers underground safety, hazardous materials handling, and emergency rescue procedures."
        PackType = "Department"
        IsActive = $true
        PolicyCount = 8
        IsSequential = $true
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Remote Work & Hybrid Policy Pack"
        PackDescription = "Policies governing remote and hybrid working arrangements. Covers home office requirements, data security at home, communication expectations, productivity tracking, and ergonomic self-assessment."
        PackType = "Custom"
        IsActive = $true
        PolicyCount = 4
        IsSequential = $false
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Procurement & Supply Chain Pack"
        PackDescription = "Policies for procurement and supply chain staff. Covers vendor due diligence, B-BBEE preferential procurement, tender process compliance, conflict of interest, and ethical sourcing."
        PackType = "Role"
        IsActive = $true
        PolicyCount = 5
        IsSequential = $false
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Executive Leadership Governance Pack"
        PackDescription = "Governance and fiduciary policies for C-suite and senior leadership. Covers King IV compliance, board reporting obligations, insider trading, and executive code of ethics."
        PackType = "Role"
        IsActive = $true
        PolicyCount = 6
        IsSequential = $false
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Driver & Fleet Safety Pack"
        PackDescription = "Policies for employees who operate company vehicles or drive for work purposes. Covers road safety, vehicle inspection protocols, accident reporting, and National Road Traffic Act compliance."
        PackType = "Role"
        IsActive = $true
        PolicyCount = 4
        IsSequential = $true
        TargetProcessType = "Joiner"
    },
    @{
        PackName = "Customer-Facing Staff Pack"
        PackDescription = "Policies for frontline and customer-facing employees. Covers customer interaction standards, complaints handling (CPA compliance), data collection consent, and brand representation guidelines."
        PackType = "Role"
        IsActive = $true
        PolicyCount = 5
        IsSequential = $false
        TargetProcessType = "Joiner"
    }
)

# ============================================================================
# Seed the list
# ============================================================================

Write-Host "Seeding $($policyPacks.Count) policy packs into $listName..." -ForegroundColor Yellow
Write-Host ""

$created = 0
$skipped = 0

foreach ($pack in $policyPacks) {
    # Check if pack already exists
    $existing = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='PackName'/><Value Type='Text'>$($pack.PackName)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue

    if ($existing) {
        Write-Host "  Exists: $($pack.PackName)" -ForegroundColor Gray
        $skipped++
        continue
    }

    # Get random policy IDs based on pack's policy count
    $policyIds = Get-RandomPolicyIds -Count $pack.PolicyCount

    $values = @{
        "Title"             = $pack.PackName
        "PackName"          = $pack.PackName
        "PackDescription"   = $pack.PackDescription
        "PackType"          = $pack.PackType
        "IsActive"          = $pack.IsActive
        "PolicyIds"         = $policyIds
        "PolicyCount"       = $pack.PolicyCount
        "IsSequential"      = $pack.IsSequential
        "TargetProcessType" = $pack.TargetProcessType
    }

    try {
        Add-PnPListItem -List $listName -Values $values | Out-Null
        Write-Host "  Created: $($pack.PackName) ($($pack.PackType), $($pack.PolicyCount) policies)" -ForegroundColor Green
        $created++
    } catch {
        Write-Host "  FAILED: $($pack.PackName) - $_" -ForegroundColor Red
    }
}

# ============================================================================
# Summary
# ============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Seeding Complete" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Gray
Write-Host "  Created: $created" -ForegroundColor Green
Write-Host "  Skipped: $skipped (already existed)" -ForegroundColor Gray
Write-Host "  Total packs in script: $($policyPacks.Count)" -ForegroundColor White
Write-Host ""
Write-Host "  Pack Types:" -ForegroundColor White
Write-Host "    Onboarding:  2  (New Employee, Contractor)" -ForegroundColor Gray
Write-Host "    Department:  3  (IT, H&S, Financial, Mining)" -ForegroundColor Gray
Write-Host "    Role:        4  (Management, Procurement, Executive, Driver, Customer)" -ForegroundColor Gray
Write-Host "    Location:    3  (JHB, CPT, DBN)" -ForegroundColor Gray
Write-Host "    Custom:      6  (POPIA, B-BBEE, Exit, Transfer, Annual, Remote, etc.)" -ForegroundColor Gray
Write-Host ""
Write-Host "  Process Types:" -ForegroundColor White
Write-Host "    Joiner:  17" -ForegroundColor Gray
Write-Host "    Mover:   1  (Internal Transfer)" -ForegroundColor Gray
Write-Host "    Leaver:  1  (Exit & Offboarding)" -ForegroundColor Gray
Write-Host ""
Write-Host "  Site: $SiteUrl" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan

Disconnect-PnPOnline
