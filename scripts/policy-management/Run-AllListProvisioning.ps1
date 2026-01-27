# ============================================================================
# DWx Policy Manager - Run All List Provisioning (No Connect)
# Master script that runs all list provisioning - assumes already connected
# ============================================================================
#
# USAGE:
#   First connect to SharePoint:
#     Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive
#   Then run this script:
#     .\Run-AllListProvisioning.ps1
#
# ============================================================================

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  DWx Policy Manager - List Provisioning" -ForegroundColor Cyan
Write-Host "  Running all list creation scripts (no authentication)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# Verify connection
try {
    $web = Get-PnPWeb -ErrorAction Stop
    Write-Host "`nConnected to: $($web.Title)" -ForegroundColor Green
    Write-Host "URL: $($web.Url)" -ForegroundColor Gray
} catch {
    Write-Host "`nERROR: Not connected to SharePoint!" -ForegroundColor Red
    Write-Host "Please connect first using:" -ForegroundColor Yellow
    Write-Host '  Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive' -ForegroundColor White
    exit 1
}

# ============================================================================
# CORE LISTS (1-3): PM_Policies, PM_PolicyVersions, PM_PolicyAcknowledgements
# ============================================================================
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  PHASE 1: Creating Core Policy Lists" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan

# PM_Policies
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

Write-Host "  PM_Policies list configured" -ForegroundColor Green

# PM_PolicyVersions
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

# PM_PolicyAcknowledgements
Write-Host "`n Creating PM_PolicyAcknowledgements list..." -ForegroundColor Yellow
$listName = "PM_PolicyAcknowledgements"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Version Number" -InternalName "PolicyVersionNumber" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "AckUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Department" -InternalName "UserDepartment" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Role" -InternalName "UserRole" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Status" -InternalName "AckStatus" -Type Choice -Choices "Not Sent","Sent","Opened","In Progress","Acknowledged","Overdue","Exempted","Failed" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Assigned Date" -InternalName "AssignedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Due Date" -InternalName "DueDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "First Opened Date" -InternalName "FirstOpenedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledged Date" -InternalName "AcknowledgedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Document Open Count" -InternalName "DocumentOpenCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Read Time (Seconds)" -InternalName "TotalReadTimeSeconds" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Last Accessed Date" -InternalName "LastAccessedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Device Type" -InternalName "DeviceType" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledgement Text" -InternalName "AcknowledgementText" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Digital Signature" -InternalName "DigitalSignature" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledgement Method" -InternalName "AcknowledgementMethod" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Required" -InternalName "QuizRequired" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Status" -InternalName "QuizStatus" -Type Choice -Choices "Not Started","In Progress","Passed","Failed","Exempted" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Score" -InternalName "QuizScore" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Attempts" -InternalName "QuizAttempts" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Completed Date" -InternalName "QuizCompletedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Delegated" -InternalName "IsDelegated" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Delegated By" -InternalName "DelegatedBy" -Type User -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Delegation Reason" -InternalName "DelegationReason" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Reminders Sent" -InternalName "RemindersSent" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Last Reminder Date" -InternalName "LastReminderDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Manager Notified" -InternalName "ManagerNotified" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Exempted" -InternalName "IsExempted" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Exemption ID" -InternalName "ExemptionId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Compliant" -InternalName "IsCompliant" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Overdue Days" -InternalName "OverdueDays" -Type Number -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "AckStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "UserEmail" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyAcknowledgements list configured" -ForegroundColor Green

Write-Host "`n  Core lists created successfully!" -ForegroundColor Green

# ============================================================================
# QUIZ LISTS (4-6): PM_PolicyQuizzes, PM_PolicyQuizQuestions, PM_PolicyQuizResults
# ============================================================================
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  PHASE 2: Creating Quiz Lists" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan

# PM_PolicyQuizzes
Write-Host "`n Creating PM_PolicyQuizzes list..." -ForegroundColor Yellow
$listName = "PM_PolicyQuizzes"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Title" -InternalName "QuizTitle" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Description" -InternalName "QuizDescription" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Passing Score" -InternalName "PassingScore" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Allow Retake" -InternalName "AllowRetake" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Max Attempts" -InternalName "MaxAttempts" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Time Limit (Minutes)" -InternalName "TimeLimit" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Randomize Questions" -InternalName "RandomizeQuestions" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Show Correct Answers" -InternalName "ShowCorrectAnswers" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyQuizzes list configured" -ForegroundColor Green

# PM_PolicyQuizQuestions
Write-Host "`n Creating PM_PolicyQuizQuestions list..." -ForegroundColor Yellow
$listName = "PM_PolicyQuizQuestions"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Quiz ID" -InternalName "QuizId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Question Text" -InternalName "QuestionText" -Type Note -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Question Type" -InternalName "QuestionType" -Type Choice -Choices "MultipleChoice","TrueFalse","MultiSelect","ShortAnswer" -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Options" -InternalName "Options" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Correct Answer" -InternalName "CorrectAnswer" -Type Note -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Points" -InternalName "Points" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Explanation" -InternalName "Explanation" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Order Index" -InternalName "OrderIndex" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Mandatory" -InternalName "IsMandatory" -Type Boolean -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "QuizId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyQuizQuestions list configured" -ForegroundColor Green

# PM_PolicyQuizResults
Write-Host "`n Creating PM_PolicyQuizResults list..." -ForegroundColor Yellow
$listName = "PM_PolicyQuizResults"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Quiz ID" -InternalName "QuizId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledgement ID" -InternalName "AcknowledgementId" -Type Number -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "QuizUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Attempt Number" -InternalName "AttemptNumber" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Score" -InternalName "Score" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Percentage" -InternalName "Percentage" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Passed" -InternalName "Passed" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Started Date" -InternalName "StartedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Completed Date" -InternalName "CompletedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Time Spent (Seconds)" -InternalName "TimeSpentSeconds" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Answers" -InternalName "Answers" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Correct Answers" -InternalName "CorrectAnswers" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Incorrect Answers" -InternalName "IncorrectAnswers" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Skipped Questions" -InternalName "SkippedQuestions" -Type Number -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "QuizId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "AcknowledgementId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyQuizResults list configured" -ForegroundColor Green

Write-Host "`n  Quiz lists created successfully!" -ForegroundColor Green

# ============================================================================
# WORKFLOW LISTS (7-9): PM_PolicyExemptions, PM_PolicyDistributions, PM_PolicyTemplates
# ============================================================================
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  PHASE 3: Creating Workflow Lists" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan

# PM_PolicyExemptions
Write-Host "`n Creating PM_PolicyExemptions list..." -ForegroundColor Yellow
$listName = "PM_PolicyExemptions"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "ExemptUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Exemption Reason" -InternalName "ExemptionReason" -Type Note -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Exemption Type" -InternalName "ExemptionType" -Type Choice -Choices "Temporary","Permanent","Conditional" -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Status" -InternalName "ExemptionStatus" -Type Choice -Choices "Pending","Approved","Denied","Expired","Revoked" -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Request Date" -InternalName "RequestDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Effective Date" -InternalName "EffectiveDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Expiry Date" -InternalName "ExpiryDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Requested By" -InternalName "RequestedBy" -Type User -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Reviewed By" -InternalName "ReviewedBy" -Type User -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Reviewed Date" -InternalName "ReviewedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Review Comments" -InternalName "ReviewComments" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approved By" -InternalName "ApprovedBy" -Type User -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approved Date" -InternalName "ApprovedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Compensating Controls" -InternalName "CompensatingControls" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Alternative Requirements" -InternalName "AlternativeRequirements" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Revoked By" -InternalName "RevokedBy" -Type User -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Revoked Date" -InternalName "RevokedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Revoked Reason" -InternalName "RevokedReason" -Type Note -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "ExemptionStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyExemptions list configured" -ForegroundColor Green

# PM_PolicyDistributions
Write-Host "`n Creating PM_PolicyDistributions list..." -ForegroundColor Yellow
$listName = "PM_PolicyDistributions"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Distribution Name" -InternalName "DistributionName" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Distribution Scope" -InternalName "DistributionScope" -Type Choice -Choices "All Employees","Department","Location","Role","Custom","New Hires Only" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Scheduled Date" -InternalName "ScheduledDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Distributed Date" -InternalName "DistributedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Target Count" -InternalName "TargetCount" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Sent" -InternalName "TotalSent" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Delivered" -InternalName "TotalDelivered" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Opened" -InternalName "TotalOpened" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Acknowledged" -InternalName "TotalAcknowledged" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Overdue" -InternalName "TotalOverdue" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Exempted" -InternalName "TotalExempted" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Failed" -InternalName "TotalFailed" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Due Date" -InternalName "DueDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Reminder Schedule" -InternalName "ReminderSchedule" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Escalation Enabled" -InternalName "EscalationEnabled" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Completed Date" -InternalName "CompletedDate" -Type DateTime -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyDistributions list configured" -ForegroundColor Green

# PM_PolicyTemplates
Write-Host "`n Creating PM_PolicyTemplates list..." -ForegroundColor Yellow
$listName = "PM_PolicyTemplates"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Template Name" -InternalName "TemplateName" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Template Category" -InternalName "TemplateCategory" -Type Choice -Choices "HR Policies","IT & Security","Health & Safety","Compliance","Financial","Operational","Legal","Environmental","Quality Assurance","Data Privacy","Custom" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Description" -InternalName "TemplateDescription" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "HTML Template" -InternalName "HTMLTemplate" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Document Template URL" -InternalName "DocumentTemplateURL" -Type URL -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Default Acknowledgement Type" -InternalName "DefaultAcknowledgementType" -Type Choice -Choices "One-Time","Periodic - Annual","Periodic - Quarterly","Periodic - Monthly","On Update","Conditional" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Default Deadline (Days)" -InternalName "DefaultDeadlineDays" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Default Requires Quiz" -InternalName "DefaultRequiresQuiz" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Default Review Cycle (Months)" -InternalName "DefaultReviewCycleMonths" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Default Compliance Risk" -InternalName "DefaultComplianceRisk" -Type Choice -Choices "Critical","High","Medium","Low","Informational" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Usage Count" -InternalName "UsageCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyTemplates list configured" -ForegroundColor Green

Write-Host "`n  Workflow lists created successfully!" -ForegroundColor Green

# ============================================================================
# SOCIAL LISTS (10-14): PM_PolicyRatings, PM_PolicyComments, etc.
# ============================================================================
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  PHASE 4: Creating Social Lists" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan

# PM_PolicyRatings
Write-Host "`n Creating PM_PolicyRatings list..." -ForegroundColor Yellow
$listName = "PM_PolicyRatings"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "RatingUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Rating" -InternalName "Rating" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Rating Date" -InternalName "RatingDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Review Title" -InternalName "ReviewTitle" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Review Text" -InternalName "ReviewText" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Review Helpful Count" -InternalName "ReviewHelpfulCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Verified Reader" -InternalName "IsVerifiedReader" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Role" -InternalName "UserRole" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Department" -InternalName "UserDepartment" -Type Text -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyRatings list configured" -ForegroundColor Green

# PM_PolicyComments
Write-Host "`n Creating PM_PolicyComments list..." -ForegroundColor Yellow
$listName = "PM_PolicyComments"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "CommentUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Comment Text" -InternalName "CommentText" -Type Note -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Comment Date" -InternalName "CommentDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Modified Date" -InternalName "CommentModifiedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Edited" -InternalName "IsEdited" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Parent Comment ID" -InternalName "ParentCommentId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Reply Count" -InternalName "ReplyCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Like Count" -InternalName "LikeCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Staff Response" -InternalName "IsStaffResponse" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Approved" -InternalName "IsApproved" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Deleted" -InternalName "IsDeleted" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Deleted Reason" -InternalName "DeletedReason" -Type Text -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "ParentCommentId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyComments list configured" -ForegroundColor Green

# PM_PolicyCommentLikes
Write-Host "`n Creating PM_PolicyCommentLikes list..." -ForegroundColor Yellow
$listName = "PM_PolicyCommentLikes"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Comment ID" -InternalName "CommentId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "LikeUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Liked Date" -InternalName "LikedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "CommentId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyCommentLikes list configured" -ForegroundColor Green

# PM_PolicyShares
Write-Host "`n Creating PM_PolicyShares list..." -ForegroundColor Yellow
$listName = "PM_PolicyShares"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Shared By" -InternalName "SharedBy" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Shared By Email" -InternalName "SharedByEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Share Method" -InternalName "ShareMethod" -Type Choice -Choices "Email","Teams","Link","QRCode","Download" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Share Date" -InternalName "ShareDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Share Message" -InternalName "ShareMessage" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Shared With Emails" -InternalName "SharedWithEmails" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Teams Channel ID" -InternalName "SharedWithTeamsChannelId" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "View Count" -InternalName "ViewCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "First Viewed Date" -InternalName "FirstViewedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Last Viewed Date" -InternalName "LastViewedDate" -Type DateTime -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyShares list configured" -ForegroundColor Green

# PM_PolicyFollowers
Write-Host "`n Creating PM_PolicyFollowers list..." -ForegroundColor Yellow
$listName = "PM_PolicyFollowers"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "FollowerUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Followed Date" -InternalName "FollowedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Notify On Update" -InternalName "NotifyOnUpdate" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Notify On Comment" -InternalName "NotifyOnComment" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Notify On New Version" -InternalName "NotifyOnNewVersion" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Email Notifications" -InternalName "EmailNotifications" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Teams Notifications" -InternalName "TeamsNotifications" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "In-App Notifications" -InternalName "InAppNotifications" -Type Boolean -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "UserEmail" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyFollowers list configured" -ForegroundColor Green

Write-Host "`n  Social lists created successfully!" -ForegroundColor Green

# ============================================================================
# POLICY PACK LISTS (15-16): PM_PolicyPacks, PM_PolicyPackAssignments
# ============================================================================
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  PHASE 5: Creating Policy Pack Lists" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan

# PM_PolicyPacks
Write-Host "`n Creating PM_PolicyPacks list..." -ForegroundColor Yellow
$listName = "PM_PolicyPacks"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Pack Name" -InternalName "PackName" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Pack Description" -InternalName "PackDescription" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Pack Category" -InternalName "PackCategory" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Pack Type" -InternalName "PackType" -Type Choice -Choices "Onboarding","Department","Role","Location","Custom" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Mandatory" -InternalName "IsMandatory" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Target Departments" -InternalName "TargetDepartments" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Target Roles" -InternalName "TargetRoles" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Target Locations" -InternalName "TargetLocations" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy IDs" -InternalName "PolicyIds" -Type Note -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Count" -InternalName "PolicyCount" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Require All Acknowledged" -InternalName "RequireAllAcknowledged" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledgement Deadline (Days)" -InternalName "AcknowledgementDeadlineDays" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Read Timeframe" -InternalName "ReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6","Custom" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Sequential" -InternalName "IsSequential" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Sequence" -InternalName "PolicySequence" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Send Welcome Email" -InternalName "SendWelcomeEmail" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Send Teams Notification" -InternalName "SendTeamsNotification" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Welcome Email Template" -InternalName "WelcomeEmailTemplate" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Teams Message Template" -InternalName "TeamsMessageTemplate" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Assignments" -InternalName "TotalAssignments" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Completed" -InternalName "TotalCompleted" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Average Completion Days" -InternalName "AverageCompletionDays" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Completion Rate" -InternalName "CompletionRate" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Version" -InternalName "PackVersion" -Type Text -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PackType" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyPacks list configured" -ForegroundColor Green

# PM_PolicyPackAssignments
Write-Host "`n Creating PM_PolicyPackAssignments list..." -ForegroundColor Yellow
$listName = "PM_PolicyPackAssignments"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Pack ID" -InternalName "PackId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "AssignedUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Department" -InternalName "UserDepartment" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Role" -InternalName "UserRole" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Assigned Date" -InternalName "AssignedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Assigned By" -InternalName "AssignedBy" -Type User -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Assignment Reason" -InternalName "AssignmentReason" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Onboarding Stage" -InternalName "OnboardingStage" -Type Choice -Choices "Pre-Start","Day 1","Week 1","Month 1","Month 3" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Due Date" -InternalName "DueDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Read Timeframe" -InternalName "ReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6","Custom" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Policies" -InternalName "TotalPolicies" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledged Policies" -InternalName "AcknowledgedPolicies" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Pending Policies" -InternalName "PendingPolicies" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Overdue Policies" -InternalName "OverduePolicies" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Progress Percentage" -InternalName "ProgressPercentage" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Status" -InternalName "AssignmentStatus" -Type Choice -Choices "Not Started","In Progress","Completed","Overdue","Exempted" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Started Date" -InternalName "StartedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Completed Date" -InternalName "CompletedDate" -Type DateTime -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PackId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "UserEmail" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "AssignmentStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyPackAssignments list configured" -ForegroundColor Green

Write-Host "`n  Policy Pack lists created successfully!" -ForegroundColor Green

# ============================================================================
# ANALYTICS LISTS (17-20): PM_PolicyAuditLog, PM_PolicyAnalytics, etc.
# ============================================================================
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  PHASE 6: Creating Analytics & Audit Lists" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan

# PM_PolicyAuditLog
Write-Host "`n Creating PM_PolicyAuditLog list..." -ForegroundColor Yellow
$listName = "PM_PolicyAuditLog"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Entity Type" -InternalName "EntityType" -Type Choice -Choices "Policy","Acknowledgement","Exemption","Distribution","Quiz","Template" -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Entity ID" -InternalName "EntityId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Action" -InternalName "AuditAction" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Action Description" -InternalName "ActionDescription" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Performed By" -InternalName "PerformedBy" -Type User -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Performed By Email" -InternalName "PerformedByEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "IP Address" -InternalName "IPAddress" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Agent" -InternalName "UserAgent" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Device Type" -InternalName "DeviceType" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Old Value" -InternalName "OldValue" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "New Value" -InternalName "NewValue" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Change Details" -InternalName "ChangeDetails" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Action Date" -InternalName "ActionDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Compliance Relevant" -InternalName "ComplianceRelevant" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Regulatory Impact" -InternalName "RegulatoryImpact" -Type Text -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "EntityType" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "AuditAction" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyAuditLog list configured" -ForegroundColor Green

# PM_PolicyAnalytics
Write-Host "`n Creating PM_PolicyAnalytics list..." -ForegroundColor Yellow
$listName = "PM_PolicyAnalytics"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Analytics Date" -InternalName "AnalyticsDate" -Type DateTime -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Period Type" -InternalName "PeriodType" -Type Choice -Choices "Daily","Weekly","Monthly","Quarterly" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Views" -InternalName "TotalViews" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Unique Viewers" -InternalName "UniqueViewers" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Average Read Time (Seconds)" -InternalName "AverageReadTimeSeconds" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Downloads" -InternalName "TotalDownloads" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Assigned" -InternalName "TotalAssigned" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Acknowledged" -InternalName "TotalAcknowledged" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Overdue" -InternalName "TotalOverdue" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Compliance Rate" -InternalName "ComplianceRate" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Average Time to Acknowledge (Days)" -InternalName "AverageTimeToAcknowledgeDays" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Quiz Attempts" -InternalName "TotalQuizAttempts" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Average Quiz Score" -InternalName "AverageQuizScore" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Pass Rate" -InternalName "QuizPassRate" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Feedback" -InternalName "TotalFeedback" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Positive Feedback" -InternalName "PositiveFeedback" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Negative Feedback" -InternalName "NegativeFeedback" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "High Risk Non-Compliance" -InternalName "HighRiskNonCompliance" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Escalated Cases" -InternalName "EscalatedCases" -Type Number -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "AnalyticsDate" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PeriodType" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyAnalytics list configured" -ForegroundColor Green

# PM_PolicyFeedback
Write-Host "`n Creating PM_PolicyFeedback list..." -ForegroundColor Yellow
$listName = "PM_PolicyFeedback"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "FeedbackUser" -Type User -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Feedback Type" -InternalName "FeedbackType" -Type Choice -Choices "Question","Suggestion","Issue","Compliment" -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Feedback Text" -InternalName "FeedbackText" -Type Note -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Anonymous" -InternalName "IsAnonymous" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Response Text" -InternalName "ResponseText" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Responded By" -InternalName "RespondedBy" -Type User -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Responded Date" -InternalName "RespondedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Status" -InternalName "FeedbackStatus" -Type Choice -Choices "Open","InProgress","Resolved","Closed" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Priority" -InternalName "FeedbackPriority" -Type Choice -Choices "Critical","High","Medium","Low" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Public" -InternalName "IsPublic" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Helpful Count" -InternalName "HelpfulCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Submitted Date" -InternalName "SubmittedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Resolved Date" -InternalName "ResolvedDate" -Type DateTime -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "FeedbackStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyFeedback list configured" -ForegroundColor Green

# PM_PolicyDocuments
Write-Host "`n Creating PM_PolicyDocuments list..." -ForegroundColor Yellow
$listName = "PM_PolicyDocuments"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Document Type" -InternalName "DocumentType" -Type Choice -Choices "Primary","Appendix","Form","Template","Guide","Reference" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Document Category" -InternalName "DocumentCategory" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "File Name" -InternalName "FileName" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "File URL" -InternalName "FileURL" -Type URL -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "File Size" -InternalName "FileSize" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "File Extension" -InternalName "FileExtension" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Document Title" -InternalName "DocumentTitle" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Document Description" -InternalName "DocumentDescription" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Document Keywords" -InternalName "DocumentKeywords" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Document Version" -InternalName "DocumentVersion" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Current Version" -InternalName "IsCurrentVersion" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Security Classification" -InternalName "SecurityClassification" -Type Choice -Choices "Public","Internal","Confidential","Restricted" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "View Count" -InternalName "ViewCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Download Count" -InternalName "DownloadCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Featured" -InternalName "IsFeatured" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "DocumentType" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "  PM_PolicyDocuments list configured" -ForegroundColor Green

Write-Host "`n  Analytics & Audit lists created successfully!" -ForegroundColor Green

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  DWx POLICY MANAGER - LIST PROVISIONING COMPLETE!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  20 Lists Created:" -ForegroundColor White
Write-Host "  ─────────────────────────────────────────────────────────" -ForegroundColor Gray
Write-Host "  Core (3):       PM_Policies, PM_PolicyVersions, PM_PolicyAcknowledgements" -ForegroundColor White
Write-Host "  Quiz (3):       PM_PolicyQuizzes, PM_PolicyQuizQuestions, PM_PolicyQuizResults" -ForegroundColor White
Write-Host "  Workflow (3):   PM_PolicyExemptions, PM_PolicyDistributions, PM_PolicyTemplates" -ForegroundColor White
Write-Host "  Social (5):     PM_PolicyRatings, PM_PolicyComments, PM_PolicyCommentLikes," -ForegroundColor White
Write-Host "                  PM_PolicyShares, PM_PolicyFollowers" -ForegroundColor White
Write-Host "  Packs (2):      PM_PolicyPacks, PM_PolicyPackAssignments" -ForegroundColor White
Write-Host "  Analytics (4):  PM_PolicyAuditLog, PM_PolicyAnalytics, PM_PolicyFeedback," -ForegroundColor White
Write-Host "                  PM_PolicyDocuments" -ForegroundColor White
Write-Host ""
Write-Host "  All lists use PM_ prefix (Policy Manager)" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Green
