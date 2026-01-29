# ============================================================================
# Create-PolicyManagementLists.ps1
# Creates all SharePoint lists for the Policy Management System
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,

    [Parameter(Mandatory=$false)]
    [switch]$IncludeSampleData
)

# Import PnP PowerShell
Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host "JML Policy Management System - SharePoint List Provisioning" -ForegroundColor Cyan
Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host ""

# Connect to SharePoint
Write-Host "Connecting to SharePoint site: $SiteUrl" -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Host "Connected successfully!`n" -ForegroundColor Green

# ============================================================================
# 1. PM_Policies List
# ============================================================================

Write-Host "Creating PM_Policies list..." -ForegroundColor Yellow

$policiesList = Get-PnPList -Identity "PM_Policies" -ErrorAction SilentlyContinue
if ($null -eq $policiesList) {
    New-PnPList -Title "PM_Policies" -Template GenericList -OnQuickLaunch

    # Basic Information
    Add-PnPField -List "PM_Policies" -DisplayName "PolicyNumber" -InternalName "PolicyNumber" -Type Text -Required -AddToDefaultView
    Add-PnPField -List "PM_Policies" -DisplayName "PolicyName" -InternalName "PolicyName" -Type Text -Required -AddToDefaultView
    Add-PnPField -List "PM_Policies" -DisplayName "PolicyCategory" -InternalName "PolicyCategory" -Type Choice -Choices "HR Policies","IT & Security","Health & Safety","Compliance","Financial","Operational","Legal","Environmental","Quality Assurance","Data Privacy","Custom" -AddToDefaultView
    Add-PnPField -List "PM_Policies" -DisplayName "PolicyType" -InternalName "PolicyType" -Type Choice -Choices "Corporate","Departmental","Regional","Role-Specific","Project-Specific","Regulatory"
    Add-PnPField -List "PM_Policies" -DisplayName "Description" -InternalName "Description" -Type Note

    # Version Management
    Add-PnPField -List "PM_Policies" -DisplayName "VersionNumber" -InternalName "VersionNumber" -Type Text -AddToDefaultView
    Add-PnPField -List "PM_Policies" -DisplayName "VersionType" -InternalName "VersionType" -Type Choice -Choices "Major","Minor","Draft"
    Add-PnPField -List "PM_Policies" -DisplayName "MajorVersion" -InternalName "MajorVersion" -Type Number
    Add-PnPField -List "PM_Policies" -DisplayName "MinorVersion" -InternalName "MinorVersion" -Type Number

    # Document
    Add-PnPField -List "PM_Policies" -DisplayName "DocumentFormat" -InternalName "DocumentFormat" -Type Choice -Choices "PDF","Word","HTML","Markdown","External Link"
    Add-PnPField -List "PM_Policies" -DisplayName "DocumentURL" -InternalName "DocumentURL" -Type URL
    Add-PnPField -List "PM_Policies" -DisplayName "HTMLContent" -InternalName "HTMLContent" -Type Note

    # Ownership
    Add-PnPField -List "PM_Policies" -DisplayName "PolicyOwner" -InternalName "PolicyOwner" -Type User -AddToDefaultView
    Add-PnPField -List "PM_Policies" -DisplayName "DepartmentOwner" -InternalName "DepartmentOwner" -Type Text

    # Status & Lifecycle
    Add-PnPField -List "PM_Policies" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Draft","In Review","Pending Approval","Approved","Published","Archived","Retired","Expired" -AddToDefaultView
    Add-PnPField -List "PM_Policies" -DisplayName "EffectiveDate" -InternalName "EffectiveDate" -Type DateTime
    Add-PnPField -List "PM_Policies" -DisplayName "ExpiryDate" -InternalName "ExpiryDate" -Type DateTime
    Add-PnPField -List "PM_Policies" -DisplayName "NextReviewDate" -InternalName "NextReviewDate" -Type DateTime
    Add-PnPField -List "PM_Policies" -DisplayName "ReviewCycleMonths" -InternalName "ReviewCycleMonths" -Type Number
    Add-PnPField -List "PM_Policies" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -AddToDefaultView
    Add-PnPField -List "PM_Policies" -DisplayName "IsMandatory" -InternalName "IsMandatory" -Type Boolean

    # Classification
    Add-PnPField -List "PM_Policies" -DisplayName "Tags" -InternalName "Tags" -Type Note
    Add-PnPField -List "PM_Policies" -DisplayName "RelatedPolicyIds" -InternalName "RelatedPolicyIds" -Type Note
    Add-PnPField -List "PM_Policies" -DisplayName "SupersedesPolicyId" -InternalName "SupersedesPolicyId" -Type Number
    Add-PnPField -List "PM_Policies" -DisplayName "RegulatoryReference" -InternalName "RegulatoryReference" -Type Text
    Add-PnPField -List "PM_Policies" -DisplayName "ComplianceRisk" -InternalName "ComplianceRisk" -Type Choice -Choices "Critical","High","Medium","Low","Informational" -AddToDefaultView

    # Acknowledgement Configuration
    Add-PnPField -List "PM_Policies" -DisplayName "RequiresAcknowledgement" -InternalName "RequiresAcknowledgement" -Type Boolean
    Add-PnPField -List "PM_Policies" -DisplayName "AcknowledgementType" -InternalName "AcknowledgementType" -Type Choice -Choices "One-Time","Periodic - Annual","Periodic - Quarterly","Periodic - Monthly","On Update","Conditional"
    Add-PnPField -List "PM_Policies" -DisplayName "AcknowledgementDeadlineDays" -InternalName "AcknowledgementDeadlineDays" -Type Number
    Add-PnPField -List "PM_Policies" -DisplayName "ReadTimeframe" -InternalName "ReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6","Custom" -AddToDefaultView
    Add-PnPField -List "PM_Policies" -DisplayName "ReadTimeframeDays" -InternalName "ReadTimeframeDays" -Type Number
    Add-PnPField -List "PM_Policies" -DisplayName "RequiresQuiz" -InternalName "RequiresQuiz" -Type Boolean
    Add-PnPField -List "PM_Policies" -DisplayName "QuizPassingScore" -InternalName "QuizPassingScore" -Type Number
    Add-PnPField -List "PM_Policies" -DisplayName "AllowRetake" -InternalName "AllowRetake" -Type Boolean
    Add-PnPField -List "PM_Policies" -DisplayName "MaxRetakeAttempts" -InternalName "MaxRetakeAttempts" -Type Number

    # Distribution
    Add-PnPField -List "PM_Policies" -DisplayName "DistributionScope" -InternalName "DistributionScope" -Type Choice -Choices "All Employees","Department","Location","Role","Custom","New Hires Only"
    Add-PnPField -List "PM_Policies" -DisplayName "TargetDepartments" -InternalName "TargetDepartments" -Type Note
    Add-PnPField -List "PM_Policies" -DisplayName "TargetLocations" -InternalName "TargetLocations" -Type Note
    Add-PnPField -List "PM_Policies" -DisplayName "TargetRoles" -InternalName "TargetRoles" -Type Note

    # Analytics
    Add-PnPField -List "PM_Policies" -DisplayName "TotalDistributed" -InternalName "TotalDistributed" -Type Number
    Add-PnPField -List "PM_Policies" -DisplayName "TotalAcknowledged" -InternalName "TotalAcknowledged" -Type Number
    Add-PnPField -List "PM_Policies" -DisplayName "CompliancePercentage" -InternalName "CompliancePercentage" -Type Number
    Add-PnPField -List "PM_Policies" -DisplayName "AverageReadTime" -InternalName "AverageReadTime" -Type Number
    Add-PnPField -List "PM_Policies" -DisplayName "AverageTimeToAcknowledge" -InternalName "AverageTimeToAcknowledge" -Type Number

    # Metadata
    Add-PnPField -List "PM_Policies" -DisplayName "Keywords" -InternalName "Keywords" -Type Note
    Add-PnPField -List "PM_Policies" -DisplayName "Language" -InternalName "Language" -Type Text
    Add-PnPField -List "PM_Policies" -DisplayName "ReadabilityScore" -InternalName "ReadabilityScore" -Type Number
    Add-PnPField -List "PM_Policies" -DisplayName "EstimatedReadTimeMinutes" -InternalName "EstimatedReadTimeMinutes" -Type Number

    # Workflow Dates
    Add-PnPField -List "PM_Policies" -DisplayName "SubmittedForReviewDate" -InternalName "SubmittedForReviewDate" -Type DateTime
    Add-PnPField -List "PM_Policies" -DisplayName "ReviewCompletedDate" -InternalName "ReviewCompletedDate" -Type DateTime
    Add-PnPField -List "PM_Policies" -DisplayName "ApprovedDate" -InternalName "ApprovedDate" -Type DateTime
    Add-PnPField -List "PM_Policies" -DisplayName "PublishedDate" -InternalName "PublishedDate" -Type DateTime
    Add-PnPField -List "PM_Policies" -DisplayName "ArchivedDate" -InternalName "ArchivedDate" -Type DateTime

    # Additional
    Add-PnPField -List "PM_Policies" -DisplayName "Comments" -InternalName "Comments" -Type Note
    Add-PnPField -List "PM_Policies" -DisplayName "InternalNotes" -InternalName "InternalNotes" -Type Note
    Add-PnPField -List "PM_Policies" -DisplayName "PublicComments" -InternalName "PublicComments" -Type Note

    Write-Host "✓ PM_Policies list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_Policies list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 2. PM_PolicyVersions List
# ============================================================================

Write-Host "Creating PM_PolicyVersions list..." -ForegroundColor Yellow

$versionsLis = Get-PnPList -Identity "PM_PolicyVersions" -ErrorAction SilentlyContinue
if ($null -eq $versionsList) {
    New-PnPList -Title "PM_PolicyVersions" -Template GenericList -OnQuickLaunch

    Add-PnPField -List "PM_PolicyVersions" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyVersions" -DisplayName "VersionNumber" -InternalName "VersionNumber" -Type Text -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyVersions" -DisplayName "VersionType" -InternalName "VersionType" -Type Choice -Choices "Major","Minor","Draft" -AddToDefaultView
    Add-PnPField -List "PM_PolicyVersions" -DisplayName "ChangeDescription" -InternalName "ChangeDescription" -Type Note -Required
    Add-PnPField -List "PM_PolicyVersions" -DisplayName "ChangeSummary" -InternalName "ChangeSummary" -Type Note
    Add-PnPField -List "PM_PolicyVersions" -DisplayName "DocumentURL" -InternalName "DocumentURL" -Type URL
    Add-PnPField -List "PM_PolicyVersions" -DisplayName "HTMLContent" -InternalName "HTMLContent" -Type Note
    Add-PnPField -List "PM_PolicyVersions" -DisplayName "EffectiveDate" -InternalName "EffectiveDate" -Type DateTime -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyVersions" -DisplayName "IsCurrentVersion" -InternalName "IsCurrentVersion" -Type Boolean -AddToDefaultView

    Write-Host "✓ PM_PolicyVersions list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyVersions list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 3. PM_PolicyAcknowledgements List
# ============================================================================

Write-Host "Creating PM_PolicyAcknowledgements list..." -ForegroundColor Yellow

$ackList = Get-PnPList -Identity "PM_PolicyAcknowledgements" -ErrorAction SilentlyContinue
if ($null -eq $ackList) {
    New-PnPList -Title "PM_PolicyAcknowledgements" -Template GenericList -OnQuickLaunch

    # Policy & User
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "PolicyVersionNumber" -InternalName "PolicyVersionNumber" -Type Text -Required
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "UserId" -InternalName "UserId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "UserEmail" -InternalName "UserEmail" -Type Text
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "UserDepartment" -InternalName "UserDepartment" -Type Text
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "UserRole" -InternalName "UserRole" -Type Text
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "UserLocation" -InternalName "UserLocation" -Type Text

    # Status & Tracking
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Not Sent","Sent","Opened","In Progress","Acknowledged","Overdue","Exempted","Failed" -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "AssignedDate" -InternalName "AssignedDate" -Type DateTime -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "DueDate" -InternalName "DueDate" -Type DateTime -AddToDefaultView
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "FirstOpenedDate" -InternalName "FirstOpenedDate" -Type DateTime
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "AcknowledgedDate" -InternalName "AcknowledgedDate" -Type DateTime -AddToDefaultView

    # Reading Analytics
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "DocumentOpenCount" -InternalName "DocumentOpenCount" -Type Number
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "TotalReadTimeSeconds" -InternalName "TotalReadTimeSeconds" -Type Number
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "LastAccessedDate" -InternalName "LastAccessedDate" -Type DateTime
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "IPAddress" -InternalName "IPAddress" -Type Text
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "DeviceType" -InternalName "DeviceType" -Type Text

    # Acknowledgement Details
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "AcknowledgementText" -InternalName "AcknowledgementText" -Type Note
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "DigitalSignature" -InternalName "DigitalSignature" -Type Note
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "AcknowledgementMethod" -InternalName "AcknowledgementMethod" -Type Text
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "PhotoEvidenceURL" -InternalName "PhotoEvidenceURL" -Type URL

    # Quiz Results
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "QuizRequired" -InternalName "QuizRequired" -Type Boolean
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "QuizStatus" -InternalName "QuizStatus" -Type Choice -Choices "Not Started","In Progress","Passed","Failed","Exempted"
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "QuizScore" -InternalName "QuizScore" -Type Number
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "QuizAttempts" -InternalName "QuizAttempts" -Type Number
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "QuizCompletedDate" -InternalName "QuizCompletedDate" -Type DateTime

    # Delegation
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "IsDelegated" -InternalName "IsDelegated" -Type Boolean
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "DelegatedById" -InternalName "DelegatedById" -Type Number
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "DelegationReason" -InternalName "DelegationReason" -Type Note

    # Reminders
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "RemindersSent" -InternalName "RemindersSent" -Type Number
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "LastReminderDate" -InternalName "LastReminderDate" -Type DateTime
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "EscalationLevel" -InternalName "EscalationLevel" -Type Number
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "ManagerNotified" -InternalName "ManagerNotified" -Type Boolean
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "ManagerNotifiedDate" -InternalName "ManagerNotifiedDate" -Type DateTime

    # Exemptions & Compliance
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "IsExempted" -InternalName "IsExempted" -Type Boolean
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "ExemptionId" -InternalName "ExemptionId" -Type Number
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "IsCompliant" -InternalName "IsCompliant" -Type Boolean -AddToDefaultView
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "ComplianceDate" -InternalName "ComplianceDate" -Type DateTime
    Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "OverdueDays" -InternalName "OverdueDays" -Type Number

    Write-Host "✓ PM_PolicyAcknowledgements list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyAcknowledgements list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 4. PM_PolicyExemptions List
# ============================================================================

Write-Host "Creating PM_PolicyExemptions list..." -ForegroundColor Yellow

$exemptionsList = Get-PnPList -Identity "PM_PolicyExemptions" -ErrorAction SilentlyContinue
if ($null -eq $exemptionsList) {
    New-PnPList -Title "PM_PolicyExemptions" -Template GenericList -OnQuickLaunch

    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "UserId" -InternalName "UserId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "ExemptionReason" -InternalName "ExemptionReason" -Type Note -Required
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "ExemptionType" -InternalName "ExemptionType" -Type Choice -Choices "Temporary","Permanent","Conditional" -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Pending","Approved","Denied","Expired","Revoked" -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "RequestDate" -InternalName "RequestDate" -Type DateTime -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "EffectiveDate" -InternalName "EffectiveDate" -Type DateTime
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "ExpiryDate" -InternalName "ExpiryDate" -Type DateTime
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "RequestedById" -InternalName "RequestedById" -Type Number
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "ReviewedById" -InternalName "ReviewedById" -Type Number
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "ReviewedDate" -InternalName "ReviewedDate" -Type DateTime
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "ReviewComments" -InternalName "ReviewComments" -Type Note
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "ApprovedById" -InternalName "ApprovedById" -Type Number
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "ApprovedDate" -InternalName "ApprovedDate" -Type DateTime
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "CompensatingControls" -InternalName "CompensatingControls" -Type Note
    Add-PnPField -List "PM_PolicyExemptions" -DisplayName "AlternativeRequirements" -InternalName "AlternativeRequirements" -Type Note

    Write-Host "✓ PM_PolicyExemptions list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyExemptions list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 5. PM_PolicyDistributions List
# ============================================================================

Write-Host "Creating PM_PolicyDistributions list..." -ForegroundColor Yellow

$distributionsList = Get-PnPList -Identity "PM_PolicyDistributions" -ErrorAction SilentlyContinue
if ($null -eq $distributionsList) {
    New-PnPList -Title "PM_PolicyDistributions" -Template GenericList -OnQuickLaunch

    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "DistributionName" -InternalName "DistributionName" -Type Text -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "DistributionScope" -InternalName "DistributionScope" -Type Choice -Choices "All Employees","Department","Location","Role","Custom","New Hires Only" -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "ScheduledDate" -InternalName "ScheduledDate" -Type DateTime
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "DistributedDate" -InternalName "DistributedDate" -Type DateTime -AddToDefaultView
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "TargetCount" -InternalName "TargetCount" -Type Number -AddToDefaultView
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "TotalSent" -InternalName "TotalSent" -Type Number
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "TotalDelivered" -InternalName "TotalDelivered" -Type Number
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "TotalOpened" -InternalName "TotalOpened" -Type Number
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "TotalAcknowledged" -InternalName "TotalAcknowledged" -Type Number
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "TotalOverdue" -InternalName "TotalOverdue" -Type Number
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "TotalExempted" -InternalName "TotalExempted" -Type Number
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "TotalFailed" -InternalName "TotalFailed" -Type Number
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "DueDate" -InternalName "DueDate" -Type DateTime
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -AddToDefaultView
    Add-PnPField -List "PM_PolicyDistributions" -DisplayName "CompletedDate" -InternalName "CompletedDate" -Type DateTime

    Write-Host "✓ PM_PolicyDistributions list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyDistributions list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 6. PM_PolicyTemplates List
# ============================================================================

Write-Host "Creating PM_PolicyTemplates list..." -ForegroundColor Yellow

$templatesList = Get-PnPList -Identity "PM_PolicyTemplates" -ErrorAction SilentlyContinue
if ($null -eq $templatesList) {
    New-PnPList -Title "PM_PolicyTemplates" -Template GenericList -OnQuickLaunch

    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "TemplateName" -InternalName "TemplateName" -Type Text -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "TemplateCategory" -InternalName "TemplateCategory" -Type Choice -Choices "HR Policies","IT & Security","Health & Safety","Compliance","Financial","Operational","Legal","Environmental","Quality Assurance","Data Privacy","Custom" -AddToDefaultView
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "Description" -InternalName "Description" -Type Note
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "HTMLTemplate" -InternalName "HTMLTemplate" -Type Note
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "DocumentTemplateURL" -InternalName "DocumentTemplateURL" -Type URL
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "DefaultAcknowledgementType" -InternalName "DefaultAcknowledgementType" -Type Choice -Choices "One-Time","Periodic - Annual","Periodic - Quarterly","Periodic - Monthly","On Update","Conditional"
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "DefaultDeadlineDays" -InternalName "DefaultDeadlineDays" -Type Number
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "DefaultRequiresQuiz" -InternalName "DefaultRequiresQuiz" -Type Boolean
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "DefaultReviewCycleMonths" -InternalName "DefaultReviewCycleMonths" -Type Number
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "DefaultComplianceRisk" -InternalName "DefaultComplianceRisk" -Type Choice -Choices "Critical","High","Medium","Low","Informational"
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "UsageCount" -InternalName "UsageCount" -Type Number
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -AddToDefaultView

    Write-Host "✓ PM_PolicyTemplates list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyTemplates list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 7. PM_PolicyFeedback List
# ============================================================================

Write-Host "Creating PM_PolicyFeedback list..." -ForegroundColor Yellow

$feedbackList = Get-PnPList -Identity "PM_PolicyFeedback" -ErrorAction SilentlyContinue
if ($null -eq $feedbackList) {
    New-PnPList -Title "PM_PolicyFeedback" -Template GenericList -OnQuickLaunch

    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "UserId" -InternalName "UserId" -Type Number -Required
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "FeedbackType" -InternalName "FeedbackType" -Type Choice -Choices "Question","Suggestion","Issue","Compliment" -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "FeedbackText" -InternalName "FeedbackText" -Type Note -Required
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "IsAnonymous" -InternalName "IsAnonymous" -Type Boolean
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "ResponseText" -InternalName "ResponseText" -Type Note
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "RespondedById" -InternalName "RespondedById" -Type Number
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "RespondedDate" -InternalName "RespondedDate" -Type DateTime
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Open","InProgress","Resolved","Closed" -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices "Low","Normal","High","Critical" -AddToDefaultView
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "IsPublic" -InternalName "IsPublic" -Type Boolean
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "HelpfulCount" -InternalName "HelpfulCount" -Type Number
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "SubmittedDate" -InternalName "SubmittedDate" -Type DateTime -AddToDefaultView
    Add-PnPField -List "PM_PolicyFeedback" -DisplayName "ResolvedDate" -InternalName "ResolvedDate" -Type DateTime

    Write-Host "✓ PM_PolicyFeedback list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyFeedback list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 8. PM_PolicyAuditLog List
# ============================================================================

Write-Host "Creating PM_PolicyAuditLog list..." -ForegroundColor Yellow

$auditList = Get-PnPList -Identity "PM_PolicyAuditLog" -ErrorAction SilentlyContinue
if ($null -eq $auditList) {
    New-PnPList -Title "PM_PolicyAuditLog" -Template GenericList -OnQuickLaunch

    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "EntityType" -InternalName "EntityType" -Type Choice -Choices "Policy","Acknowledgement","Exemption","Distribution","Quiz","Template" -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "EntityId" -InternalName "EntityId" -Type Number -Required
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "Action" -InternalName "Action" -Type Text -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "ActionDescription" -InternalName "ActionDescription" -Type Note -Required
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "PerformedById" -InternalName "PerformedById" -Type Number -Required
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "PerformedByEmail" -InternalName "PerformedByEmail" -Type Text -Required
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "IPAddress" -InternalName "IPAddress" -Type Text
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "UserAgent" -InternalName "UserAgent" -Type Text
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "DeviceType" -InternalName "DeviceType" -Type Text
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "OldValue" -InternalName "OldValue" -Type Note
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "NewValue" -InternalName "NewValue" -Type Note
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "ChangeDetails" -InternalName "ChangeDetails" -Type Note
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "ActionDate" -InternalName "ActionDate" -Type DateTime -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "ComplianceRelevant" -InternalName "ComplianceRelevant" -Type Boolean
    Add-PnPField -List "PM_PolicyAuditLog" -DisplayName "RegulatoryImpact" -InternalName "RegulatoryImpact" -Type Text

    Write-Host "✓ PM_PolicyAuditLog list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyAuditLog list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 9. PM_PolicyDocuments List (Policy Hub Document Center)
# ============================================================================

Write-Host "Creating PM_PolicyDocuments list..." -ForegroundColor Yellow

$documentsList = Get-PnPList -Identity "PM_PolicyDocuments" -ErrorAction SilentlyContinue
if ($null -eq $documentsList) {
    New-PnPList -Title "PM_PolicyDocuments" -Template DocumentLibrary -OnQuickLaunch

    # Document Classification
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DocumentType" -InternalName "DocumentType" -Type Choice -Choices "Primary","Appendix","Form","Template","Guide","Reference" -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DocumentCategory" -InternalName "DocumentCategory" -Type Text -AddToDefaultView
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DocumentSubcategory" -InternalName "DocumentSubcategory" -Type Text

    # Rich Metadata
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DocumentTitle" -InternalName "DocumentTitle" -Type Text -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DocumentDescription" -InternalName "DocumentDescription" -Type Note
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DocumentSummary" -InternalName "DocumentSummary" -Type Note
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DocumentKeywords" -InternalName "DocumentKeywords" -Type Note
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DocumentAuthor" -InternalName "DocumentAuthor" -Type Text
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DocumentOwner" -InternalName "DocumentOwner" -Type User

    # Versioning
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DocumentVersion" -InternalName "DocumentVersion" -Type Text -AddToDefaultView
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DocumentVersionDate" -InternalName "DocumentVersionDate" -Type DateTime
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "IsCurrentVersion" -InternalName "IsCurrentVersion" -Type Boolean -AddToDefaultView

    # Classification & Tagging
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "SecurityClassification" -InternalName "SecurityClassification" -Type Choice -Choices "Public","Internal","Confidential","Restricted" -AddToDefaultView
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "Audience" -InternalName "Audience" -Type Note
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "Department" -InternalName "Department" -Type Note
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "Location" -InternalName "Location" -Type Note
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "Tags" -InternalName "Tags" -Type Note

    # Lifecycle
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "PublishedDate" -InternalName "PublishedDate" -Type DateTime
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "ExpiryDate" -InternalName "ExpiryDate" -Type DateTime
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "ReviewDate" -InternalName "ReviewDate" -Type DateTime

    # Access Control
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "RequiresApproval" -InternalName "RequiresApproval" -Type Boolean
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "RestrictedAccess" -InternalName "RestrictedAccess" -Type Boolean
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "AllowedRoles" -InternalName "AllowedRoles" -Type Note

    # Analytics
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "ViewCount" -InternalName "ViewCount" -Type Number
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "DownloadCount" -InternalName "DownloadCount" -Type Number
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "LastViewedDate" -InternalName "LastViewedDate" -Type DateTime
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "AverageRating" -InternalName "AverageRating" -Type Number
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "RatingCount" -InternalName "RatingCount" -Type Number

    # Search Optimization
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "SearchKeywords" -InternalName "SearchKeywords" -Type Note
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "SearchBoost" -InternalName "SearchBoost" -Type Number
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "IsFeatured" -InternalName "IsFeatured" -Type Boolean
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "IsPopular" -InternalName "IsPopular" -Type Boolean

    # Relationships
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "RelatedDocumentIds" -InternalName "RelatedDocumentIds" -Type Note
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "ParentDocumentId" -InternalName "ParentDocumentId" -Type Number

    # Status
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -AddToDefaultView
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "IsArchived" -InternalName "IsArchived" -Type Boolean
    Add-PnPField -List "PM_PolicyDocuments" -DisplayName "ArchiveReason" -InternalName "ArchiveReason" -Type Note

    Write-Host "✓ PM_PolicyDocuments library created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyDocuments library already exists" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host "List Provisioning Completed Successfully!" -ForegroundColor Green
Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "The following lists have been created:" -ForegroundColor Yellow
Write-Host "  1. PM_Policies" -ForegroundColor White
Write-Host "  2. PM_PolicyVersions" -ForegroundColor White
Write-Host "  3. PM_PolicyAcknowledgements" -ForegroundColor White
Write-Host "  4. PM_PolicyExemptions" -ForegroundColor White
Write-Host "  5. PM_PolicyDistributions" -ForegroundColor White
Write-Host "  6. PM_PolicyTemplates" -ForegroundColor White
Write-Host "  7. PM_PolicyFeedback" -ForegroundColor White
Write-Host "  8. PM_PolicyAuditLog" -ForegroundColor White
Write-Host "  9. PM_PolicyDocuments (Document Library)" -ForegroundColor White
Write-Host ""
Write-Host "Enhanced Features:" -ForegroundColor Yellow
Write-Host "  ✓ Read Timeframe Tracking (Day 1, Week 1, Month 1, etc.)" -ForegroundColor Green
Write-Host "  ✓ Policy Hub Document Center with rich metadata" -ForegroundColor Green
Write-Host "  ✓ Advanced filtering and sorting capabilities" -ForegroundColor Green
Write-Host "  ✓ Document classification and tagging" -ForegroundColor Green
Write-Host "  ✓ Analytics tracking (views, downloads, ratings)" -ForegroundColor Green
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  • Run Add-PolicySampleData.ps1 to add sample policies" -ForegroundColor White
Write-Host "  • Configure permissions for policy authors and administrators" -ForegroundColor White
Write-Host "  • Deploy the SPFx web parts for Policy Management" -ForegroundColor White
Write-Host "  • Configure read timeframes for policies" -ForegroundColor White
Write-Host "  • Upload policy documents to the Policy Hub" -ForegroundColor White
Write-Host ""

# Disconnect
# Disconnect-PnPOnline
