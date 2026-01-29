# ============================================================================
# DWx Policy Manager - Seed Current User Data
# Creates sample policies and acknowledgements for the logged-in user
# This script is designed to be run to quickly test the MyPolicies webpart
# ============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"
)

# Connect to SharePoint
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "DWx Policy Manager - Current User Data Seeding" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Connecting to SharePoint: $SiteUrl" -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -Interactive
Write-Host "Connected successfully!`n" -ForegroundColor Green

# Get current user info
$web = Get-PnPWeb
$currentUser = Get-PnPProperty -ClientObject $web -Property CurrentUser
$currentUserId = $currentUser.Id
$currentUserEmail = $currentUser.Email
$currentUserName = $currentUser.Title

Write-Host "Current User: $currentUserName ($currentUserEmail)" -ForegroundColor Gray
Write-Host "User ID: $currentUserId`n" -ForegroundColor Gray

# ============================================================================
# STEP 1: Check/Create Sample Policies
# ============================================================================
Write-Host "Step 1: Checking PM_Policies list..." -ForegroundColor Yellow

$existingPolicies = Get-PnPListItem -List "PM_Policies" -PageSize 100 -ErrorAction SilentlyContinue
$policyCount = if ($existingPolicies) { $existingPolicies.Count } else { 0 }

if ($policyCount -lt 5) {
    Write-Host "  Creating sample policies..." -ForegroundColor Gray

    $policies = @(
        @{
            Title = "POL-HR-001 Employee Code of Conduct"
            PolicyNumber = "POL-HR-001"
            PolicyName = "Employee Code of Conduct"
            PolicyCategory = "HR Policies"
            PolicyType = "Corporate"
            PolicyDescription = "Standards of professional conduct expected from all employees."
            VersionNumber = "3.0"
            PolicyStatus = "Published"
            ComplianceRisk = "Critical"
            IsMandatory = $true
            IsActive = $true
            RequiresAcknowledgement = $true
            AcknowledgementDeadlineDays = 14
            ReadTimeframe = "Week 1"
            RequiresQuiz = $true
            QuizPassingScore = 80
            EffectiveDate = (Get-Date).AddDays(-60)
        },
        @{
            Title = "POL-HR-002 Anti-Harassment Policy"
            PolicyNumber = "POL-HR-002"
            PolicyName = "Anti-Harassment and Discrimination Policy"
            PolicyCategory = "HR Policies"
            PolicyType = "Corporate"
            PolicyDescription = "Prohibits all forms of harassment and discrimination in the workplace."
            VersionNumber = "2.1"
            PolicyStatus = "Published"
            ComplianceRisk = "Critical"
            IsMandatory = $true
            IsActive = $true
            RequiresAcknowledgement = $true
            AcknowledgementDeadlineDays = 7
            ReadTimeframe = "Day 3"
            RequiresQuiz = $true
            QuizPassingScore = 85
            EffectiveDate = (Get-Date).AddDays(-90)
        },
        @{
            Title = "POL-IT-001 Information Security Policy"
            PolicyNumber = "POL-IT-001"
            PolicyName = "Information Security Policy"
            PolicyCategory = "IT & Security"
            PolicyType = "Corporate"
            PolicyDescription = "Framework for protecting company information assets and systems."
            VersionNumber = "2.0"
            PolicyStatus = "Published"
            ComplianceRisk = "High"
            IsMandatory = $true
            IsActive = $true
            RequiresAcknowledgement = $true
            AcknowledgementDeadlineDays = 14
            ReadTimeframe = "Week 1"
            RequiresQuiz = $true
            QuizPassingScore = 80
            EffectiveDate = (Get-Date).AddDays(-45)
        },
        @{
            Title = "POL-COM-001 Data Privacy Policy"
            PolicyNumber = "POL-COM-001"
            PolicyName = "Data Privacy and Protection Policy"
            PolicyCategory = "Compliance"
            PolicyType = "Corporate"
            PolicyDescription = "GDPR and data protection requirements for handling personal data."
            VersionNumber = "1.5"
            PolicyStatus = "Published"
            ComplianceRisk = "High"
            IsMandatory = $true
            IsActive = $true
            RequiresAcknowledgement = $true
            AcknowledgementDeadlineDays = 7
            ReadTimeframe = "Day 3"
            RequiresQuiz = $false
            EffectiveDate = (Get-Date).AddDays(-30)
        },
        @{
            Title = "POL-HS-001 Health and Safety Policy"
            PolicyNumber = "POL-HS-001"
            PolicyName = "Workplace Health and Safety Policy"
            PolicyCategory = "Health & Safety"
            PolicyType = "Corporate"
            PolicyDescription = "Ensures a safe and healthy working environment for all employees."
            VersionNumber = "4.0"
            PolicyStatus = "Published"
            ComplianceRisk = "Medium"
            IsMandatory = $true
            IsActive = $true
            RequiresAcknowledgement = $true
            AcknowledgementDeadlineDays = 14
            ReadTimeframe = "Week 2"
            RequiresQuiz = $false
            EffectiveDate = (Get-Date).AddDays(-120)
        }
    )

    foreach ($policy in $policies) {
        try {
            Add-PnPListItem -List "PM_Policies" -Values $policy | Out-Null
            Write-Host "    Created: $($policy.PolicyNumber)" -ForegroundColor Green
        } catch {
            Write-Host "    Exists/Skipped: $($policy.PolicyNumber)" -ForegroundColor Gray
        }
    }

    # Refresh policy list
    $existingPolicies = Get-PnPListItem -List "PM_Policies" -PageSize 100 -ErrorAction SilentlyContinue
    $policyCount = $existingPolicies.Count
}

Write-Host "  Total policies in list: $policyCount`n" -ForegroundColor Green

# ============================================================================
# STEP 2: Create Acknowledgements for Current User
# ============================================================================
Write-Host "Step 2: Creating acknowledgements for current user..." -ForegroundColor Yellow

# Get policy IDs
$policyIds = $existingPolicies | Select-Object -First 10 | ForEach-Object { $_.Id }

if ($policyIds.Count -eq 0) {
    Write-Host "  ERROR: No policies found in PM_Policies list!" -ForegroundColor Red
    Write-Host "  Please run the policy provisioning scripts first." -ForegroundColor Yellow
    exit
}

# Check existing acknowledgements for this user
$existingAcks = Get-PnPListItem -List "PM_PolicyAcknowledgements" -Query "<View><Query><Where><Eq><FieldRef Name='AckUserId'/><Value Type='Number'>$currentUserId</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue

if ($existingAcks -and $existingAcks.Count -gt 0) {
    Write-Host "  User already has $($existingAcks.Count) acknowledgements" -ForegroundColor Gray
    Write-Host "  Skipping acknowledgement creation`n" -ForegroundColor Gray
} else {
    Write-Host "  Creating acknowledgements for $($policyIds.Count) policies..." -ForegroundColor Gray

    $statuses = @("Sent", "Sent", "Opened", "In Progress", "Acknowledged", "Acknowledged", "Overdue")
    $index = 0

    foreach ($policyId in $policyIds) {
        $status = $statuses[$index % $statuses.Count]
        $assignedDate = (Get-Date).AddDays(-($index * 3 + 5))
        $dueDate = $assignedDate.AddDays(14)

        $ackValues = @{
            "Title" = "ACK-$policyId-$currentUserId"
            "PolicyId" = $policyId
            "PolicyVersionNumber" = "1.0"
            "AckUserId" = $currentUserId
            "UserEmail" = $currentUserEmail
            "UserDepartment" = "General"
            "AckStatus" = $status
            "AssignedDate" = $assignedDate
            "DueDate" = $dueDate
        }

        # Add acknowledged details for completed policies
        if ($status -eq "Acknowledged") {
            $ackValues["AcknowledgedDate"] = $assignedDate.AddDays((Get-Random -Minimum 1 -Maximum 10))
            $ackValues["ReadDuration"] = Get-Random -Minimum 180 -Maximum 900
        }

        try {
            Add-PnPListItem -List "PM_PolicyAcknowledgements" -Values $ackValues | Out-Null
            Write-Host "    Created acknowledgement for Policy ID: $policyId (Status: $status)" -ForegroundColor Green
        } catch {
            Write-Host "    Failed: Policy ID $policyId - $_" -ForegroundColor Red
        }

        $index++
    }
}

# ============================================================================
# STEP 3: Create Policy Packs (optional)
# ============================================================================
Write-Host "`nStep 3: Checking PM_PolicyPacks list..." -ForegroundColor Yellow

$existingPacks = Get-PnPListItem -List "PM_PolicyPacks" -PageSize 100 -ErrorAction SilentlyContinue
$packCount = if ($existingPacks) { $existingPacks.Count } else { 0 }

if ($packCount -eq 0) {
    Write-Host "  Creating sample policy pack..." -ForegroundColor Gray

    # Get first 3 policy IDs for the pack
    $packPolicyIds = $policyIds | Select-Object -First 3
    $packPolicyIdsJson = ConvertTo-Json @($packPolicyIds) -Compress

    $packValues = @{
        "Title" = "New Employee Onboarding Pack"
        "PackName" = "New Employee Onboarding Pack"
        "PackDescription" = "Essential policies for new employees to acknowledge during their first week"
        "PackCategory" = "Onboarding"
        "PackType" = "Onboarding"
        "IsActive" = $true
        "IsMandatory" = $true
        "PolicyIds" = $packPolicyIdsJson
        "PolicyCount" = $packPolicyIds.Count
        "AcknowledgementDeadlineDays" = 7
        "ReadTimeframe" = "Week 1"
    }

    try {
        $pack = Add-PnPListItem -List "PM_PolicyPacks" -Values $packValues
        Write-Host "    Created pack: New Employee Onboarding Pack`n" -ForegroundColor Green

        # Create pack assignment for current user
        Write-Host "  Creating pack assignment for current user..." -ForegroundColor Gray

        $assignmentValues = @{
            "Title" = "Pack Assignment - $currentUserName"
            "PackId" = $pack.Id
            "UserId" = $currentUserId
            "UserEmail" = $currentUserEmail
            "AssignedDate" = (Get-Date).AddDays(-5)
            "DueDate" = (Get-Date).AddDays(7)
            "Status" = "In Progress"
            "CompletedPolicies" = 0
            "TotalPolicies" = $packPolicyIds.Count
        }

        Add-PnPListItem -List "PM_PolicyPackAssignments" -Values $assignmentValues | Out-Null
        Write-Host "    Created pack assignment`n" -ForegroundColor Green
    } catch {
        Write-Host "    Failed to create pack: $_" -ForegroundColor Red
    }
} else {
    Write-Host "  Policy packs already exist ($packCount found)`n" -ForegroundColor Gray
}

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SEEDING COMPLETE!" -ForegroundColor Green
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Summary for user: $currentUserName" -ForegroundColor White
Write-Host "  - Policies in system: $policyCount" -ForegroundColor Gray

$finalAcks = Get-PnPListItem -List "PM_PolicyAcknowledgements" -Query "<View><Query><Where><Eq><FieldRef Name='AckUserId'/><Value Type='Number'>$currentUserId</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
$ackCount = if ($finalAcks) { $finalAcks.Count } else { 0 }
Write-Host "  - Your acknowledgements: $ackCount" -ForegroundColor Gray

$finalPackAssignments = Get-PnPListItem -List "PM_PolicyPackAssignments" -Query "<View><Query><Where><Eq><FieldRef Name='UserId'/><Value Type='Number'>$currentUserId</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
$packAssignmentCount = if ($finalPackAssignments) { $finalPackAssignments.Count } else { 0 }
Write-Host "  - Your pack assignments: $packAssignmentCount" -ForegroundColor Gray

Write-Host "`nNext Steps:" -ForegroundColor Yellow
Write-Host "  1. Deploy the updated SPFx package to SharePoint" -ForegroundColor White
Write-Host "  2. Navigate to the PolicyManager site" -ForegroundColor White
Write-Host "  3. Add the MyPolicies or PolicyHub webpart to a page" -ForegroundColor White

Disconnect-PnPOnline
Write-Host "`nDone!" -ForegroundColor Green
