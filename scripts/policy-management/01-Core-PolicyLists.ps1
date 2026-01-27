# ============================================================================
# DWx Policy Manager - Core Lists Provisioning
# Part 1: PM_Policies, PM_PolicyVersions, PM_PolicyAcknowledgements
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,

    [Parameter(Mandatory=$false)]
    [switch]$UseWebLogin
)

# Connect to SharePoint
Write-Host "Connecting to SharePoint: $SiteUrl" -ForegroundColor Cyan
if ($UseWebLogin) {
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
} else {
    Connect-PnPOnline -Url $SiteUrl -Interactive
}

Write-Host "Connected successfully!" -ForegroundColor Green

# ============================================================================
# LIST 1: PM_Policies (Master Policy List)
# ============================================================================
Write-Host "`n Creating PM_Policies list..." -ForegroundColor Yellow

$listName = "PM_Policies"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

# Basic Information Fields
Add-PnPField -List $listName -DisplayName "Policy Number" -InternalName "PolicyNumber" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Name" -InternalName "PolicyName" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Category" -InternalName "PolicyCategory" -Type Choice -Choices "HR Policies","IT & Security","Health & Safety","Compliance","Financial","Operational","Legal","Environmental","Quality Assurance","Data Privacy","Custom" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Type" -InternalName "PolicyType" -Type Choice -Choices "Corporate","Departmental","Regional","Role-Specific","Project-Specific","Regulatory" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Description" -InternalName "PolicyDescription" -Type Note -ErrorAction SilentlyContinue

# Version Management Fields
Add-PnPField -List $listName -DisplayName "Version Number" -InternalName "VersionNumber" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Version Type" -InternalName "VersionType" -Type Choice -Choices "Major","Minor","Draft" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Major Version" -InternalName "MajorVersion" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Minor Version" -InternalName "MinorVersion" -Type Number -ErrorAction SilentlyContinue

# Document Fields
Add-PnPField -List $listName -DisplayName "Document Format" -InternalName "DocumentFormat" -Type Choice -Choices "PDF","Word","HTML","Markdown","External Link" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Document URL" -InternalName "DocumentURL" -Type URL -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "HTML Content" -InternalName "HTMLContent" -Type Note -ErrorAction SilentlyContinue

# Ownership Fields
Add-PnPField -List $listName -DisplayName "Policy Owner" -InternalName "PolicyOwner" -Type User -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Authors" -InternalName "PolicyAuthors" -Type UserMulti -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Department Owner" -InternalName "DepartmentOwner" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Reviewers" -InternalName "Reviewers" -Type UserMulti -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approvers" -InternalName "Approvers" -Type UserMulti -ErrorAction SilentlyContinue

# Status & Lifecycle Fields
Add-PnPField -List $listName -DisplayName "Status" -InternalName "PolicyStatus" -Type Choice -Choices "Draft","In Review","Pending Approval","Approved","Published","Archived","Retired","Expired" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Effective Date" -InternalName "EffectiveDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Expiry Date" -InternalName "ExpiryDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Next Review Date" -InternalName "NextReviewDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Review Cycle (Months)" -InternalName "ReviewCycleMonths" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Mandatory" -InternalName "IsMandatory" -Type Boolean -ErrorAction SilentlyContinue

# Classification Fields
Add-PnPField -List $listName -DisplayName "Tags" -InternalName "Tags" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Related Policy IDs" -InternalName "RelatedPolicyIds" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Supersedes Policy ID" -InternalName "SupersedesPolicyId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Regulatory Reference" -InternalName "RegulatoryReference" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Compliance Risk" -InternalName "ComplianceRisk" -Type Choice -Choices "Critical","High","Medium","Low","Informational" -AddToDefaultView -ErrorAction SilentlyContinue

# Acknowledgement Configuration Fields
Add-PnPField -List $listName -DisplayName "Requires Acknowledgement" -InternalName "RequiresAcknowledgement" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledgement Type" -InternalName "AcknowledgementType" -Type Choice -Choices "One-Time","Periodic - Annual","Periodic - Quarterly","Periodic - Monthly","On Update","Conditional" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledgement Deadline (Days)" -InternalName "AcknowledgementDeadlineDays" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Read Timeframe" -InternalName "ReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6","Custom" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Read Timeframe Days" -InternalName "ReadTimeframeDays" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Requires Quiz" -InternalName "RequiresQuiz" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Passing Score" -InternalName "QuizPassingScore" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Allow Retake" -InternalName "AllowRetake" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Max Retake Attempts" -InternalName "MaxRetakeAttempts" -Type Number -ErrorAction SilentlyContinue

# Distribution Fields
Add-PnPField -List $listName -DisplayName "Distribution Scope" -InternalName "DistributionScope" -Type Choice -Choices "All Employees","Department","Location","Role","Custom","New Hires Only" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Target Departments" -InternalName "TargetDepartments" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Target Locations" -InternalName "TargetLocations" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Target Roles" -InternalName "TargetRoles" -Type Note -ErrorAction SilentlyContinue

# Analytics Fields
Add-PnPField -List $listName -DisplayName "Total Distributed" -InternalName "TotalDistributed" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Acknowledged" -InternalName "TotalAcknowledged" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Compliance Percentage" -InternalName "CompliancePercentage" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Average Read Time" -InternalName "AverageReadTime" -Type Number -ErrorAction SilentlyContinue

# Workflow Date Fields
Add-PnPField -List $listName -DisplayName "Submitted for Review Date" -InternalName "SubmittedForReviewDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Review Completed Date" -InternalName "ReviewCompletedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approved Date" -InternalName "ApprovedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Published Date" -InternalName "PublishedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Archived Date" -InternalName "ArchivedDate" -Type DateTime -ErrorAction SilentlyContinue

# Additional Fields
Add-PnPField -List $listName -DisplayName "Internal Notes" -InternalName "InternalNotes" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Estimated Read Time (Minutes)" -InternalName "EstimatedReadTimeMinutes" -Type Number -ErrorAction SilentlyContinue

# Create indexes for performance
Set-PnPField -List $listName -Identity "PolicyStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyCategory" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "ComplianceRisk" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_Policies list configured with all fields" -ForegroundColor Green

# ============================================================================
# LIST 2: PM_PolicyVersions (Version History)
# ============================================================================
Write-Host "`n Creating PM_PolicyVersions list..." -ForegroundColor Yellow

$listName = "PM_PolicyVersions"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Version Number" -InternalName "VersionNumber" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Version Type" -InternalName "VersionType" -Type Choice -Choices "Major","Minor","Draft" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Change Description" -InternalName "ChangeDescription" -Type Note -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Change Summary" -InternalName "ChangeSummary" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Document URL" -InternalName "DocumentURL" -Type URL -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "HTML Content" -InternalName "HTMLContent" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Effective Date" -InternalName "EffectiveDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Current Version" -InternalName "IsCurrentVersion" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Comparison URL" -InternalName "ComparisonWithPreviousURL" -Type URL -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyVersions list configured" -ForegroundColor Green

# ============================================================================
# LIST 3: PM_PolicyAcknowledgements (User Acknowledgements)
# ============================================================================
Write-Host "`n Creating PM_PolicyAcknowledgements list..." -ForegroundColor Yellow

$listName = "PM_PolicyAcknowledgements"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

# Core Fields
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Version Number" -InternalName "PolicyVersionNumber" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "AckUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Department" -InternalName "UserDepartment" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Role" -InternalName "UserRole" -Type Text -ErrorAction SilentlyContinue

# Status & Tracking
Add-PnPField -List $listName -DisplayName "Status" -InternalName "AckStatus" -Type Choice -Choices "Not Sent","Sent","Opened","In Progress","Acknowledged","Overdue","Exempted","Failed" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Assigned Date" -InternalName "AssignedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Due Date" -InternalName "DueDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "First Opened Date" -InternalName "FirstOpenedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledged Date" -InternalName "AcknowledgedDate" -Type DateTime -ErrorAction SilentlyContinue

# Reading Analytics
Add-PnPField -List $listName -DisplayName "Document Open Count" -InternalName "DocumentOpenCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Read Time (Seconds)" -InternalName "TotalReadTimeSeconds" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Last Accessed Date" -InternalName "LastAccessedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Device Type" -InternalName "DeviceType" -Type Text -ErrorAction SilentlyContinue

# Acknowledgement Details
Add-PnPField -List $listName -DisplayName "Acknowledgement Text" -InternalName "AcknowledgementText" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Digital Signature" -InternalName "DigitalSignature" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledgement Method" -InternalName "AcknowledgementMethod" -Type Text -ErrorAction SilentlyContinue

# Quiz Fields
Add-PnPField -List $listName -DisplayName "Quiz Required" -InternalName "QuizRequired" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Status" -InternalName "QuizStatus" -Type Choice -Choices "Not Started","In Progress","Passed","Failed","Exempted" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Score" -InternalName "QuizScore" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Attempts" -InternalName "QuizAttempts" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Completed Date" -InternalName "QuizCompletedDate" -Type DateTime -ErrorAction SilentlyContinue

# Delegation Fields
Add-PnPField -List $listName -DisplayName "Is Delegated" -InternalName "IsDelegated" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Delegated By" -InternalName "DelegatedBy" -Type User -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Delegation Reason" -InternalName "DelegationReason" -Type Text -ErrorAction SilentlyContinue

# Reminders
Add-PnPField -List $listName -DisplayName "Reminders Sent" -InternalName "RemindersSent" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Last Reminder Date" -InternalName "LastReminderDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Manager Notified" -InternalName "ManagerNotified" -Type Boolean -ErrorAction SilentlyContinue

# Exemption Fields
Add-PnPField -List $listName -DisplayName "Is Exempted" -InternalName "IsExempted" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Exemption ID" -InternalName "ExemptionId" -Type Number -ErrorAction SilentlyContinue

# Compliance
Add-PnPField -List $listName -DisplayName "Is Compliant" -InternalName "IsCompliant" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Overdue Days" -InternalName "OverdueDays" -Type Number -ErrorAction SilentlyContinue

# Create indexes
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "AckStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "UserEmail" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyAcknowledgements list configured" -ForegroundColor Green

Write-Host "`n✅ Core Policy lists created successfully!" -ForegroundColor Green
Write-Host "   - PM_Policies" -ForegroundColor White
Write-Host "   - PM_PolicyVersions" -ForegroundColor White
Write-Host "   - PM_PolicyAcknowledgements" -ForegroundColor White
