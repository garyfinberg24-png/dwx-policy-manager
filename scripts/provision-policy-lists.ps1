# ============================================================================
# DWx Policy Manager — SharePoint List & Library Provisioning Script
# ============================================================================
# Prerequisites:
#   Install-Module -Name PnP.PowerShell -Scope CurrentUser
#   (or) Install-Module -Name SharePointPnPPowerShellOnline -Scope CurrentUser
#
# Usage:
#   .\provision-policy-lists.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/PolicyManager"
#
# This script creates all SharePoint lists and document libraries required by
# the Policy Manager application. It is idempotent — running it again will
# skip lists that already exist.
# ============================================================================

param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [switch]$SkipConnection,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeTestData
)

$ErrorActionPreference = "Stop"

# ============================================================================
# CONNECT
# ============================================================================
if (-not $SkipConnection) {
    Write-Host "`n=== Connecting to $SiteUrl ===" -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -Interactive
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Ensure-List {
    param(
        [string]$ListName,
        [string]$Template = "GenericList",
        [string]$Description = ""
    )
    $existing = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  [EXISTS] $ListName" -ForegroundColor Yellow
        return $existing
    }
    Write-Host "  [CREATE] $ListName ($Template)" -ForegroundColor Green
    if ($Template -eq "DocumentLibrary") {
        return New-PnPList -Title $ListName -Template DocumentLibrary -Url $ListName -ErrorAction Stop
    } else {
        return New-PnPList -Title $ListName -Template GenericList -Url "Lists/$ListName" -ErrorAction Stop
    }
}

function Ensure-Field {
    param(
        [string]$ListName,
        [string]$FieldName,
        [string]$Type = "Text",
        [string]$DisplayName = "",
        [bool]$Required = $false,
        [string]$DefaultValue = "",
        [string[]]$Choices = @(),
        [string]$Group = "PM Policy Manager"
    )
    if (-not $DisplayName) { $DisplayName = $FieldName }

    $existingField = Get-PnPField -List $ListName -Identity $FieldName -ErrorAction SilentlyContinue
    if ($existingField) {
        return
    }

    $params = @{
        List         = $ListName
        InternalName = $FieldName
        DisplayName  = $DisplayName
        Type         = $Type
        Group        = $Group
        Required     = $Required
        ErrorAction  = "Stop"
    }

    switch ($Type) {
        "Choice" {
            $params.Choices = $Choices
        }
        "Note" {
            # Multi-line text — no extra params needed
        }
        "Boolean" {
            # Yes/No field
        }
        "Number" {
            # Number field
        }
        "DateTime" {
            # Date field
        }
        "User" {
            # People picker
        }
        "URL" {
            # Hyperlink
        }
        "Currency" {
            # Currency
        }
    }

    Add-PnPField @params | Out-Null

    if ($DefaultValue) {
        Set-PnPField -List $ListName -Identity $FieldName -Values @{DefaultValue = $DefaultValue }
    }
}

# ============================================================================
# 1. PM_Policies — Main Policy List
# ============================================================================
Write-Host "`n=== 1. PM_Policies (Main Policy List) ===" -ForegroundColor Cyan
Ensure-List -ListName "PM_Policies" -Description "Core policy records with metadata, ownership, and lifecycle information"

$policiesFields = @(
    # Basic Information
    @{ Name="PolicyNumber"; Type="Text"; Required=$true; DisplayName="Policy Number" }
    @{ Name="PolicyName"; Type="Text"; Required=$true; DisplayName="Policy Name" }
    @{ Name="PolicyDescription"; Type="Note"; DisplayName="Policy Description" }
    @{ Name="PolicyCategory"; Type="Choice"; DisplayName="Category"; Choices=@("HR Policies","IT & Security","Health & Safety","Compliance","Financial","Operational","Legal","Environmental","Quality Assurance","Data Privacy","Custom") }
    @{ Name="PolicyType"; Type="Choice"; DisplayName="Policy Type"; Choices=@("Corporate","Departmental","Regional","Role-Specific","Project-Specific","Regulatory") }

    # Version Management
    @{ Name="VersionNumber"; Type="Text"; DisplayName="Version Number"; Default="1.0" }
    @{ Name="VersionType"; Type="Choice"; DisplayName="Version Type"; Choices=@("Major","Minor","Draft") }
    @{ Name="MajorVersion"; Type="Number"; DisplayName="Major Version" }
    @{ Name="MinorVersion"; Type="Number"; DisplayName="Minor Version" }

    # Document
    @{ Name="DocumentFormat"; Type="Choice"; DisplayName="Document Format"; Choices=@("PDF","Word","HTML","Markdown","External Link","Excel","PowerPoint","Image") }
    @{ Name="DocumentURL"; Type="URL"; DisplayName="Document URL" }
    @{ Name="DocumentLibraryId"; Type="Number"; DisplayName="Document Library Id" }
    @{ Name="HTMLContent"; Type="Note"; DisplayName="HTML Content" }

    # Ownership
    @{ Name="PolicyOwnerId"; Type="User"; DisplayName="Policy Owner" }
    @{ Name="DepartmentOwner"; Type="Text"; DisplayName="Department Owner" }
    @{ Name="Department"; Type="Text"; DisplayName="Department" }

    # Status & Lifecycle
    @{ Name="PolicyStatus"; Type="Choice"; Required=$true; DisplayName="Status"; Choices=@("Draft","In Review","Pending Approval","Approved","Rejected","Published","Archived","Retired","Expired"); Default="Draft" }
    @{ Name="EffectiveDate"; Type="DateTime"; DisplayName="Effective Date" }
    @{ Name="ExpiryDate"; Type="DateTime"; DisplayName="Expiry Date" }
    @{ Name="NextReviewDate"; Type="DateTime"; DisplayName="Next Review Date" }
    @{ Name="ReviewDate"; Type="DateTime"; DisplayName="Review Date" }
    @{ Name="ReviewCycleMonths"; Type="Number"; DisplayName="Review Cycle (Months)" }
    @{ Name="IsActive"; Type="Boolean"; DisplayName="Is Active"; Default="1" }
    @{ Name="IsMandatory"; Type="Boolean"; DisplayName="Is Mandatory"; Default="0" }

    # Classification
    @{ Name="Tags"; Type="Note"; DisplayName="Tags (JSON)" }
    @{ Name="RelatedPolicyIds"; Type="Note"; DisplayName="Related Policy Ids (JSON)" }
    @{ Name="SupersedesPolicyId"; Type="Number"; DisplayName="Supersedes Policy Id" }
    @{ Name="RegulatoryReference"; Type="Text"; DisplayName="Regulatory Reference" }
    @{ Name="ComplianceRisk"; Type="Choice"; DisplayName="Compliance Risk"; Choices=@("Critical","High","Medium","Low","Informational"); Default="Medium" }

    # Data Classification
    @{ Name="DataClassification"; Type="Choice"; DisplayName="Data Classification"; Choices=@("Public","Internal","Confidential","Restricted","Regulated"); Default="Internal" }
    @{ Name="RetentionCategory"; Type="Choice"; DisplayName="Retention Category"; Choices=@("Standard","Extended","Regulatory","Legal","Permanent"); Default="Standard" }
    @{ Name="ContainsPII"; Type="Boolean"; DisplayName="Contains PII"; Default="0" }
    @{ Name="ContainsPHI"; Type="Boolean"; DisplayName="Contains PHI"; Default="0" }
    @{ Name="ContainsFinancialData"; Type="Boolean"; DisplayName="Contains Financial Data"; Default="0" }
    @{ Name="IsLegalHold"; Type="Boolean"; DisplayName="Is Legal Hold"; Default="0" }

    # Rating
    @{ Name="AverageRating"; Type="Number"; DisplayName="Average Rating" }
    @{ Name="RatingCount"; Type="Number"; DisplayName="Rating Count" }

    # Acknowledgement Configuration
    @{ Name="RequiresAcknowledgement"; Type="Boolean"; DisplayName="Requires Acknowledgement"; Default="1" }
    @{ Name="AcknowledgementType"; Type="Choice"; DisplayName="Acknowledgement Type"; Choices=@("One-Time","Periodic - Annual","Periodic - Quarterly","Periodic - Monthly","On Update","Conditional"); Default="One-Time" }
    @{ Name="AcknowledgementDeadlineDays"; Type="Number"; DisplayName="Ack. Deadline (Days)" }
    @{ Name="ReadTimeframe"; Type="Choice"; DisplayName="Read Timeframe"; Choices=@("Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6","Custom"); Default="Week 1" }
    @{ Name="ReadTimeframeDays"; Type="Number"; DisplayName="Read Timeframe Days" }
    @{ Name="RequiresQuiz"; Type="Boolean"; DisplayName="Requires Quiz"; Default="0" }
    @{ Name="QuizPassingScore"; Type="Number"; DisplayName="Quiz Passing Score (%)" }
    @{ Name="AllowRetake"; Type="Boolean"; DisplayName="Allow Quiz Retake"; Default="1" }
    @{ Name="MaxRetakeAttempts"; Type="Number"; DisplayName="Max Retake Attempts" }

    # Distribution
    @{ Name="DistributionScope"; Type="Choice"; DisplayName="Distribution Scope"; Choices=@("All Employees","Department","Location","Role","Custom","New Hires Only"); Default="All Employees" }
    @{ Name="TargetDepartments"; Type="Note"; DisplayName="Target Departments (JSON)" }
    @{ Name="TargetRoles"; Type="Note"; DisplayName="Target Roles (JSON)" }

    # Analytics
    @{ Name="TotalDistributed"; Type="Number"; DisplayName="Total Distributed" }
    @{ Name="TotalAcknowledged"; Type="Number"; DisplayName="Total Acknowledged" }
    @{ Name="CompliancePercentage"; Type="Number"; DisplayName="Compliance %" }
    @{ Name="AverageReadTime"; Type="Number"; DisplayName="Avg Read Time (sec)" }

    # Content
    @{ Name="PolicyContent"; Type="Note"; DisplayName="Policy Content (HTML)" }
    @{ Name="PolicySummary"; Type="Note"; DisplayName="Policy Summary" }
    @{ Name="KeyPoints"; Type="Note"; DisplayName="Key Points (JSON)" }
    @{ Name="PolicyOwner"; Type="Text"; DisplayName="Policy Owner Name" }
    @{ Name="ReviewFrequency"; Type="Text"; DisplayName="Review Frequency" }
    @{ Name="EstimatedReadTimeMinutes"; Type="Number"; DisplayName="Est. Read Time (min)" }

    # Workflow Dates
    @{ Name="PublishedDate"; Type="DateTime"; DisplayName="Published Date" }
    @{ Name="ApprovedDate"; Type="DateTime"; DisplayName="Approved Date" }
    @{ Name="ArchivedDate"; Type="DateTime"; DisplayName="Archived Date" }
    @{ Name="RejectedDate"; Type="DateTime"; DisplayName="Rejected Date" }
    @{ Name="RejectionReason"; Type="Note"; DisplayName="Rejection Reason" }
    @{ Name="PolicyVersion"; Type="Text"; DisplayName="Policy Version" }

    # Additional
    @{ Name="AttachmentURLs"; Type="Note"; DisplayName="Attachment URLs (JSON)" }
    @{ Name="Keywords"; Type="Note"; DisplayName="Keywords (JSON)" }
    @{ Name="Language"; Type="Text"; DisplayName="Language"; Default="en" }
)

foreach ($field in $policiesFields) {
    $params = @{
        ListName    = "PM_Policies"
        FieldName   = $field.Name
        Type        = $field.Type
        DisplayName = if ($field.DisplayName) { $field.DisplayName } else { $field.Name }
        Required    = if ($field.Required) { $field.Required } else { $false }
    }
    if ($field.Choices) { $params.Choices = $field.Choices }
    if ($field.Default) { $params.DefaultValue = $field.Default }
    Ensure-Field @params
}
Write-Host "  PM_Policies fields provisioned." -ForegroundColor Green


# ============================================================================
# 2. PM_PolicyDocuments — Document Library
# ============================================================================
Write-Host "`n=== 2. PM_PolicyDocuments (Document Library) ===" -ForegroundColor Cyan
Ensure-List -ListName "PM_PolicyDocuments" -Template "DocumentLibrary" -Description "Policy PDF/Word documents linked to PM_Policies records"

$docLibFields = @(
    @{ Name="PolicyId"; Type="Number"; Required=$true; DisplayName="Policy Id" }
    @{ Name="DocumentType"; Type="Choice"; DisplayName="Document Type"; Choices=@("Primary","Appendix","Form","Template","Guide","Reference"); Default="Primary" }
    @{ Name="DocumentCategory"; Type="Text"; DisplayName="Document Category" }
    @{ Name="DocumentTitle"; Type="Text"; DisplayName="Document Title" }
    @{ Name="DocumentDescription"; Type="Note"; DisplayName="Document Description" }
    @{ Name="DocumentVersion"; Type="Text"; DisplayName="Document Version"; Default="1.0" }
    @{ Name="DocumentVersionDate"; Type="DateTime"; DisplayName="Version Date" }
    @{ Name="IsCurrentVersion"; Type="Boolean"; DisplayName="Is Current Version"; Default="1" }
    @{ Name="SecurityClassification"; Type="Choice"; DisplayName="Security Classification"; Choices=@("Public","Internal","Confidential","Restricted"); Default="Internal" }
    @{ Name="RequiresApproval"; Type="Boolean"; DisplayName="Requires Approval"; Default="0" }
    @{ Name="RestrictedAccess"; Type="Boolean"; DisplayName="Restricted Access"; Default="0" }
    @{ Name="ViewCount"; Type="Number"; DisplayName="View Count" }
    @{ Name="DownloadCount"; Type="Number"; DisplayName="Download Count" }
    @{ Name="LastViewedDate"; Type="DateTime"; DisplayName="Last Viewed Date" }
    @{ Name="IsActive"; Type="Boolean"; DisplayName="Is Active"; Default="1" }
    @{ Name="IsArchived"; Type="Boolean"; DisplayName="Is Archived"; Default="0" }
    @{ Name="IsFeatured"; Type="Boolean"; DisplayName="Is Featured"; Default="0" }
    @{ Name="IsPopular"; Type="Boolean"; DisplayName="Is Popular"; Default="0" }
    @{ Name="SearchKeywords"; Type="Note"; DisplayName="Search Keywords" }
    @{ Name="Tags"; Type="Note"; DisplayName="Tags (JSON)" }
)

foreach ($field in $docLibFields) {
    $params = @{
        ListName    = "PM_PolicyDocuments"
        FieldName   = $field.Name
        Type        = $field.Type
        DisplayName = if ($field.DisplayName) { $field.DisplayName } else { $field.Name }
        Required    = if ($field.Required) { $field.Required } else { $false }
    }
    if ($field.Choices) { $params.Choices = $field.Choices }
    if ($field.Default) { $params.DefaultValue = $field.Default }
    Ensure-Field @params
}
Write-Host "  PM_PolicyDocuments fields provisioned." -ForegroundColor Green


# ============================================================================
# 3. PM_PolicyAcknowledgements — Acknowledgement Tracking
# ============================================================================
Write-Host "`n=== 3. PM_PolicyAcknowledgements ===" -ForegroundColor Cyan
Ensure-List -ListName "PM_PolicyAcknowledgements" -Description "User acknowledgement records per policy"

$ackFields = @(
    @{ Name="PolicyId"; Type="Number"; Required=$true; DisplayName="Policy Id" }
    @{ Name="PolicyVersionNumber"; Type="Text"; DisplayName="Policy Version" }
    @{ Name="AckUserId"; Type="Number"; DisplayName="User Id" }
    @{ Name="UserEmail"; Type="Text"; DisplayName="User Email" }
    @{ Name="UserDepartment"; Type="Text"; DisplayName="User Department" }
    @{ Name="AckStatus"; Type="Choice"; Required=$true; DisplayName="Status"; Choices=@("Not Sent","Sent","Opened","In Progress","Acknowledged","Overdue","Exempted","Failed"); Default="Not Sent" }
    @{ Name="AssignedDate"; Type="DateTime"; Required=$true; DisplayName="Assigned Date" }
    @{ Name="DueDate"; Type="DateTime"; DisplayName="Due Date" }
    @{ Name="FirstOpenedDate"; Type="DateTime"; DisplayName="First Opened Date" }
    @{ Name="AcknowledgedDate"; Type="DateTime"; DisplayName="Acknowledged Date" }
    @{ Name="DocumentOpenCount"; Type="Number"; DisplayName="Document Open Count" }
    @{ Name="TotalReadTimeSeconds"; Type="Number"; DisplayName="Total Read Time (sec)" }
    @{ Name="LastAccessedDate"; Type="DateTime"; DisplayName="Last Accessed Date" }
    @{ Name="DeviceType"; Type="Text"; DisplayName="Device Type" }
    @{ Name="AcknowledgementText"; Type="Note"; DisplayName="Acknowledgement Text" }
    @{ Name="DigitalSignature"; Type="Note"; DisplayName="Digital Signature" }
    @{ Name="AcknowledgementMethod"; Type="Choice"; DisplayName="Ack. Method"; Choices=@("Click","Signature","Voice","Other"); Default="Click" }
    @{ Name="QuizRequired"; Type="Boolean"; DisplayName="Quiz Required"; Default="0" }
    @{ Name="QuizId"; Type="Number"; DisplayName="Quiz Id" }
    @{ Name="QuizStatus"; Type="Choice"; DisplayName="Quiz Status"; Choices=@("Not Started","In Progress","Passed","Failed","Exempted") }
    @{ Name="QuizScore"; Type="Number"; DisplayName="Quiz Score" }
    @{ Name="QuizAttempts"; Type="Number"; DisplayName="Quiz Attempts" }
    @{ Name="QuizCompletedDate"; Type="DateTime"; DisplayName="Quiz Completed Date" }
    @{ Name="IsDelegated"; Type="Boolean"; DisplayName="Is Delegated"; Default="0" }
    @{ Name="DelegatedById"; Type="Number"; DisplayName="Delegated By Id" }
    @{ Name="RemindersSent"; Type="Number"; DisplayName="Reminders Sent" }
    @{ Name="LastReminderDate"; Type="DateTime"; DisplayName="Last Reminder Date" }
    @{ Name="EscalationLevel"; Type="Number"; DisplayName="Escalation Level" }
    @{ Name="ManagerNotified"; Type="Boolean"; DisplayName="Manager Notified"; Default="0" }
    @{ Name="IsExempted"; Type="Boolean"; DisplayName="Is Exempted"; Default="0" }
    @{ Name="IsCompliant"; Type="Boolean"; DisplayName="Is Compliant"; Default="0" }
    @{ Name="OverdueDays"; Type="Number"; DisplayName="Overdue Days" }
    @{ Name="PolicyNumber"; Type="Text"; DisplayName="Policy Number" }
    @{ Name="PolicyName"; Type="Text"; DisplayName="Policy Name" }
    @{ Name="PolicyCategory"; Type="Text"; DisplayName="Policy Category" }
    @{ Name="IsMandatory"; Type="Boolean"; DisplayName="Is Mandatory"; Default="0" }
)

foreach ($field in $ackFields) {
    $params = @{
        ListName    = "PM_PolicyAcknowledgements"
        FieldName   = $field.Name
        Type        = $field.Type
        DisplayName = if ($field.DisplayName) { $field.DisplayName } else { $field.Name }
        Required    = if ($field.Required) { $field.Required } else { $false }
    }
    if ($field.Choices) { $params.Choices = $field.Choices }
    if ($field.Default) { $params.DefaultValue = $field.Default }
    Ensure-Field @params
}
Write-Host "  PM_PolicyAcknowledgements fields provisioned." -ForegroundColor Green


# ============================================================================
# 4. PM_PolicyVersions — Version History
# ============================================================================
Write-Host "`n=== 4. PM_PolicyVersions ===" -ForegroundColor Cyan
Ensure-List -ListName "PM_PolicyVersions" -Description "Policy version history tracking"

$versionFields = @(
    @{ Name="PolicyId"; Type="Number"; Required=$true; DisplayName="Policy Id" }
    @{ Name="VersionNumber"; Type="Text"; Required=$true; DisplayName="Version Number" }
    @{ Name="VersionType"; Type="Choice"; DisplayName="Version Type"; Choices=@("Major","Minor","Draft") }
    @{ Name="ChangeDescription"; Type="Note"; Required=$true; DisplayName="Change Description" }
    @{ Name="ChangeSummary"; Type="Note"; DisplayName="Change Summary" }
    @{ Name="DocumentURL"; Type="URL"; DisplayName="Document URL" }
    @{ Name="HTMLContent"; Type="Note"; DisplayName="HTML Content" }
    @{ Name="EffectiveDate"; Type="DateTime"; DisplayName="Effective Date" }
    @{ Name="CreatedById"; Type="Number"; DisplayName="Created By Id" }
    @{ Name="IsCurrentVersion"; Type="Boolean"; DisplayName="Is Current Version"; Default="0" }
)

foreach ($field in $versionFields) {
    $params = @{
        ListName    = "PM_PolicyVersions"
        FieldName   = $field.Name
        Type        = $field.Type
        DisplayName = if ($field.DisplayName) { $field.DisplayName } else { $field.Name }
        Required    = if ($field.Required) { $field.Required } else { $false }
    }
    if ($field.Choices) { $params.Choices = $field.Choices }
    if ($field.Default) { $params.DefaultValue = $field.Default }
    Ensure-Field @params
}
Write-Host "  PM_PolicyVersions fields provisioned." -ForegroundColor Green


# ============================================================================
# 5. PM_PolicyExemptions — Exemption Requests
# ============================================================================
Write-Host "`n=== 5. PM_PolicyExemptions ===" -ForegroundColor Cyan
Ensure-List -ListName "PM_PolicyExemptions" -Description "Policy acknowledgement exemption requests"

$exemptionFields = @(
    @{ Name="PolicyId"; Type="Number"; Required=$true; DisplayName="Policy Id" }
    @{ Name="RequestedById"; Type="Number"; DisplayName="Requested By Id" }
    @{ Name="RequestedForId"; Type="Number"; DisplayName="Requested For Id" }
    @{ Name="ExemptionStatus"; Type="Choice"; DisplayName="Status"; Choices=@("Pending","Approved","Denied","Expired","Revoked"); Default="Pending" }
    @{ Name="ExemptionReason"; Type="Note"; Required=$true; DisplayName="Reason" }
    @{ Name="ApprovedById"; Type="Number"; DisplayName="Approved By Id" }
    @{ Name="ApprovalDate"; Type="DateTime"; DisplayName="Approval Date" }
    @{ Name="ExpiryDate"; Type="DateTime"; DisplayName="Expiry Date" }
    @{ Name="Notes"; Type="Note"; DisplayName="Notes" }
)

foreach ($field in $exemptionFields) {
    $params = @{
        ListName    = "PM_PolicyExemptions"
        FieldName   = $field.Name
        Type        = $field.Type
        DisplayName = if ($field.DisplayName) { $field.DisplayName } else { $field.Name }
        Required    = if ($field.Required) { $field.Required } else { $false }
    }
    if ($field.Choices) { $params.Choices = $field.Choices }
    if ($field.Default) { $params.DefaultValue = $field.Default }
    Ensure-Field @params
}
Write-Host "  PM_PolicyExemptions fields provisioned." -ForegroundColor Green


# ============================================================================
# 6. PM_PolicySourceDocuments — Staging Library (Bulk Import)
# ============================================================================
Write-Host "`n=== 6. PM_PolicySourceDocuments (Staging Library) ===" -ForegroundColor Cyan
Ensure-List -ListName "PM_PolicySourceDocuments" -Template "DocumentLibrary" -Description "Staging area for uploaded policy source documents"

$sourceDocFields = @(
    @{ Name="DocumentType"; Type="Choice"; DisplayName="Document Type"; Choices=@("Word Document","Excel Spreadsheet","PowerPoint Presentation","PDF","Image","Other") }
    @{ Name="FileStatus"; Type="Choice"; DisplayName="File Status"; Choices=@("Uploaded","Imported","Processing","Processed","Failed","Archived"); Default="Uploaded" }
    @{ Name="UploadDate"; Type="DateTime"; DisplayName="Upload Date" }
    @{ Name="ImportDate"; Type="DateTime"; DisplayName="Import Date" }
    @{ Name="PolicyCategory"; Type="Text"; DisplayName="Policy Category" }
    @{ Name="PolicyId"; Type="Number"; DisplayName="Linked Policy Id" }
    @{ Name="RequiresMetadata"; Type="Boolean"; DisplayName="Requires Metadata"; Default="1" }
    @{ Name="ExtractedContent"; Type="Note"; DisplayName="Extracted Content" }
)

foreach ($field in $sourceDocFields) {
    $params = @{
        ListName    = "PM_PolicySourceDocuments"
        FieldName   = $field.Name
        Type        = $field.Type
        DisplayName = if ($field.DisplayName) { $field.DisplayName } else { $field.Name }
        Required    = $false
    }
    if ($field.Choices) { $params.Choices = $field.Choices }
    if ($field.Default) { $params.DefaultValue = $field.Default }
    Ensure-Field @params
}
Write-Host "  PM_PolicySourceDocuments fields provisioned." -ForegroundColor Green


# ============================================================================
# 7. PM_PolicyAuditLog — Audit Trail
# ============================================================================
Write-Host "`n=== 7. PM_PolicyAuditLog ===" -ForegroundColor Cyan
Ensure-List -ListName "PM_PolicyAuditLog" -Description "Audit trail for all policy actions"

$auditFields = @(
    @{ Name="PolicyId"; Type="Number"; DisplayName="Policy Id" }
    @{ Name="Action"; Type="Choice"; DisplayName="Action"; Choices=@("Created","Updated","Published","Archived","Viewed","Downloaded","Acknowledged","QuizCompleted","Exempted","Distributed","Deleted","Restored","Approved","Rejected","Reviewed") }
    @{ Name="ActionById"; Type="Number"; DisplayName="Action By Id" }
    @{ Name="ActionByEmail"; Type="Text"; DisplayName="Action By Email" }
    @{ Name="ActionDate"; Type="DateTime"; DisplayName="Action Date" }
    @{ Name="Details"; Type="Note"; DisplayName="Details (JSON)" }
    @{ Name="IPAddress"; Type="Text"; DisplayName="IP Address" }
    @{ Name="UserAgent"; Type="Text"; DisplayName="User Agent" }
)

foreach ($field in $auditFields) {
    $params = @{
        ListName    = "PM_PolicyAuditLog"
        FieldName   = $field.Name
        Type        = $field.Type
        DisplayName = if ($field.DisplayName) { $field.DisplayName } else { $field.Name }
        Required    = $false
    }
    if ($field.Choices) { $params.Choices = $field.Choices }
    if ($field.Default) { $params.DefaultValue = $field.Default }
    Ensure-Field @params
}
Write-Host "  PM_PolicyAuditLog fields provisioned." -ForegroundColor Green


# ============================================================================
# 8. PM_PolicyReadReceipts — Read Receipt Tracking
# ============================================================================
Write-Host "`n=== 8. PM_PolicyReadReceipts ===" -ForegroundColor Cyan
Ensure-List -ListName "PM_PolicyReadReceipts" -Description "Read receipt records for compliance"

$receiptFields = @(
    @{ Name="PolicyId"; Type="Number"; Required=$true; DisplayName="Policy Id" }
    @{ Name="AcknowledgementId"; Type="Number"; DisplayName="Acknowledgement Id" }
    @{ Name="UserId"; Type="Number"; Required=$true; DisplayName="User Id" }
    @{ Name="ReceiptNumber"; Type="Text"; DisplayName="Receipt Number" }
    @{ Name="ReadDate"; Type="DateTime"; DisplayName="Read Date" }
    @{ Name="ReadDurationSeconds"; Type="Number"; DisplayName="Read Duration (sec)" }
    @{ Name="QuizScore"; Type="Number"; DisplayName="Quiz Score" }
    @{ Name="SignatureData"; Type="Note"; DisplayName="Signature Data" }
    @{ Name="CertificateURL"; Type="URL"; DisplayName="Certificate URL" }
)

foreach ($field in $receiptFields) {
    $params = @{
        ListName    = "PM_PolicyReadReceipts"
        FieldName   = $field.Name
        Type        = $field.Type
        DisplayName = if ($field.DisplayName) { $field.DisplayName } else { $field.Name }
        Required    = if ($field.Required) { $field.Required } else { $false }
    }
    if ($field.Choices) { $params.Choices = $field.Choices }
    if ($field.Default) { $params.DefaultValue = $field.Default }
    Ensure-Field @params
}
Write-Host "  PM_PolicyReadReceipts fields provisioned." -ForegroundColor Green


# ============================================================================
# 9. PM_PolicyQuizzes — Quiz Definitions
# ============================================================================
Write-Host "`n=== 9. PM_PolicyQuizzes ===" -ForegroundColor Cyan
Ensure-List -ListName "PM_PolicyQuizzes" -Description "Quiz definitions linked to policies"

$quizFields = @(
    @{ Name="PolicyId"; Type="Number"; Required=$true; DisplayName="Policy Id" }
    @{ Name="QuizTitle"; Type="Text"; DisplayName="Quiz Title" }
    @{ Name="QuizDescription"; Type="Note"; DisplayName="Quiz Description" }
    @{ Name="PassingScore"; Type="Number"; DisplayName="Passing Score (%)" }
    @{ Name="TimeLimitMinutes"; Type="Number"; DisplayName="Time Limit (min)" }
    @{ Name="MaxAttempts"; Type="Number"; DisplayName="Max Attempts" }
    @{ Name="ShuffleQuestions"; Type="Boolean"; DisplayName="Shuffle Questions"; Default="0" }
    @{ Name="ShuffleOptions"; Type="Boolean"; DisplayName="Shuffle Options"; Default="0" }
    @{ Name="ShowCorrectAnswers"; Type="Boolean"; DisplayName="Show Correct Answers"; Default="1" }
    @{ Name="IsActive"; Type="Boolean"; DisplayName="Is Active"; Default="1" }
    @{ Name="QuestionCount"; Type="Number"; DisplayName="Question Count" }
)

foreach ($field in $quizFields) {
    $params = @{
        ListName    = "PM_PolicyQuizzes"
        FieldName   = $field.Name
        Type        = $field.Type
        DisplayName = if ($field.DisplayName) { $field.DisplayName } else { $field.Name }
        Required    = if ($field.Required) { $field.Required } else { $false }
    }
    if ($field.Choices) { $params.Choices = $field.Choices }
    if ($field.Default) { $params.DefaultValue = $field.Default }
    Ensure-Field @params
}
Write-Host "  PM_PolicyQuizzes fields provisioned." -ForegroundColor Green


# ============================================================================
# 10. PM_PolicyQuizQuestions — Quiz Questions
# ============================================================================
Write-Host "`n=== 10. PM_PolicyQuizQuestions ===" -ForegroundColor Cyan
Ensure-List -ListName "PM_PolicyQuizQuestions" -Description "Individual quiz questions"

$questionFields = @(
    @{ Name="QuizId"; Type="Number"; Required=$true; DisplayName="Quiz Id" }
    @{ Name="QuestionText"; Type="Note"; Required=$true; DisplayName="Question Text" }
    @{ Name="QuestionType"; Type="Choice"; DisplayName="Question Type"; Choices=@("MultipleChoice","TrueFalse","MultiSelect","ShortAnswer"); Default="MultipleChoice" }
    @{ Name="Options"; Type="Note"; DisplayName="Options (JSON)" }
    @{ Name="CorrectAnswer"; Type="Note"; DisplayName="Correct Answer (JSON)" }
    @{ Name="Explanation"; Type="Note"; DisplayName="Explanation" }
    @{ Name="Points"; Type="Number"; DisplayName="Points"; Default="1" }
    @{ Name="SortOrder"; Type="Number"; DisplayName="Sort Order" }
    @{ Name="Difficulty"; Type="Choice"; DisplayName="Difficulty"; Choices=@("Easy","Medium","Hard") }
    @{ Name="SectionName"; Type="Text"; DisplayName="Section Name" }
)

foreach ($field in $questionFields) {
    $params = @{
        ListName    = "PM_PolicyQuizQuestions"
        FieldName   = $field.Name
        Type        = $field.Type
        DisplayName = if ($field.DisplayName) { $field.DisplayName } else { $field.Name }
        Required    = if ($field.Required) { $field.Required } else { $false }
    }
    if ($field.Choices) { $params.Choices = $field.Choices }
    if ($field.Default) { $params.DefaultValue = $field.Default }
    Ensure-Field @params
}
Write-Host "  PM_PolicyQuizQuestions fields provisioned." -ForegroundColor Green


# ============================================================================
# 11. PM_PolicyQuizResults — Quiz Attempts
# ============================================================================
Write-Host "`n=== 11. PM_PolicyQuizResults ===" -ForegroundColor Cyan
Ensure-List -ListName "PM_PolicyQuizResults" -Description "Individual quiz attempt results"

$resultFields = @(
    @{ Name="QuizId"; Type="Number"; Required=$true; DisplayName="Quiz Id" }
    @{ Name="PolicyId"; Type="Number"; DisplayName="Policy Id" }
    @{ Name="UserId"; Type="Number"; Required=$true; DisplayName="User Id" }
    @{ Name="AttemptNumber"; Type="Number"; DisplayName="Attempt Number" }
    @{ Name="Score"; Type="Number"; DisplayName="Score (%)" }
    @{ Name="Passed"; Type="Boolean"; DisplayName="Passed"; Default="0" }
    @{ Name="StartedDate"; Type="DateTime"; DisplayName="Started Date" }
    @{ Name="CompletedDate"; Type="DateTime"; DisplayName="Completed Date" }
    @{ Name="DurationSeconds"; Type="Number"; DisplayName="Duration (sec)" }
    @{ Name="Answers"; Type="Note"; DisplayName="Answers (JSON)" }
    @{ Name="CorrectCount"; Type="Number"; DisplayName="Correct Count" }
    @{ Name="TotalQuestions"; Type="Number"; DisplayName="Total Questions" }
)

foreach ($field in $resultFields) {
    $params = @{
        ListName    = "PM_PolicyQuizResults"
        FieldName   = $field.Name
        Type        = $field.Type
        DisplayName = if ($field.DisplayName) { $field.DisplayName } else { $field.Name }
        Required    = if ($field.Required) { $field.Required } else { $false }
    }
    if ($field.Choices) { $params.Choices = $field.Choices }
    if ($field.Default) { $params.DefaultValue = $field.Default }
    Ensure-Field @params
}
Write-Host "  PM_PolicyQuizResults fields provisioned." -ForegroundColor Green


# ============================================================================
# OPTIONAL: INSERT TEST DATA
# ============================================================================
if ($IncludeTestData) {
    Write-Host "`n=== Inserting Test Data ===" -ForegroundColor Cyan

    # Test Policies
    $testPolicies = @(
        @{
            Title = "Information Security Policy"
            PolicyNumber = "POL-IT-001"
            PolicyName = "Information Security Policy"
            PolicyCategory = "IT & Security"
            PolicyType = "Corporate"
            PolicyStatus = "Published"
            PolicyDescription = "This policy establishes the information security requirements for all employees to protect company assets, data, and systems."
            VersionNumber = "2.1"
            IsActive = $true
            IsMandatory = $true
            RequiresAcknowledgement = $true
            RequiresQuiz = $true
            QuizPassingScore = 80
            AllowRetake = $true
            MaxRetakeAttempts = 3
            ComplianceRisk = "High"
            DataClassification = "Confidential"
            DistributionScope = "All Employees"
            EffectiveDate = (Get-Date "2024-01-15")
            NextReviewDate = (Get-Date "2025-01-15")
            PolicyOwner = "Sarah Johnson"
            Department = "Information Technology"
            PolicySummary = "Comprehensive information security guidelines covering data protection, access controls, incident response, and acceptable use."
            EstimatedReadTimeMinutes = 15
        },
        @{
            Title = "Code of Conduct"
            PolicyNumber = "POL-HR-001"
            PolicyName = "Code of Conduct"
            PolicyCategory = "HR Policies"
            PolicyType = "Corporate"
            PolicyStatus = "Published"
            PolicyDescription = "Defines expected standards of behaviour for all employees."
            VersionNumber = "3.0"
            IsActive = $true
            IsMandatory = $true
            RequiresAcknowledgement = $true
            RequiresQuiz = $false
            ComplianceRisk = "Medium"
            DataClassification = "Internal"
            DistributionScope = "All Employees"
            EffectiveDate = (Get-Date "2023-06-01")
            NextReviewDate = (Get-Date "2025-06-01")
            PolicyOwner = "James Mitchell"
            Department = "Human Resources"
            PolicySummary = "Standards of professional conduct, workplace ethics, and employee responsibilities."
            EstimatedReadTimeMinutes = 10
        },
        @{
            Title = "Data Privacy Policy"
            PolicyNumber = "POL-COMP-001"
            PolicyName = "Data Privacy Policy"
            PolicyCategory = "Data Privacy"
            PolicyType = "Regulatory"
            PolicyStatus = "Published"
            PolicyDescription = "GDPR and POPIA compliant data privacy policy covering personal data handling, consent, and data subject rights."
            VersionNumber = "1.5"
            IsActive = $true
            IsMandatory = $true
            RequiresAcknowledgement = $true
            RequiresQuiz = $true
            QuizPassingScore = 85
            AllowRetake = $true
            MaxRetakeAttempts = 2
            ComplianceRisk = "Critical"
            DataClassification = "Restricted"
            ContainsPII = $true
            DistributionScope = "All Employees"
            EffectiveDate = (Get-Date "2024-03-01")
            NextReviewDate = (Get-Date "2025-03-01")
            PolicyOwner = "Lisa van der Berg"
            Department = "Compliance"
            PolicySummary = "Data privacy requirements under GDPR and POPIA, covering consent management, data subject rights, and breach notification."
            RegulatoryReference = "GDPR, POPIA"
            EstimatedReadTimeMinutes = 20
        },
        @{
            Title = "Health and Safety Policy"
            PolicyNumber = "POL-HS-001"
            PolicyName = "Health and Safety Policy"
            PolicyCategory = "Health & Safety"
            PolicyType = "Corporate"
            PolicyStatus = "Published"
            PolicyDescription = "Workplace health and safety requirements, emergency procedures, and incident reporting."
            VersionNumber = "4.0"
            IsActive = $true
            IsMandatory = $true
            RequiresAcknowledgement = $true
            RequiresQuiz = $true
            QuizPassingScore = 75
            AllowRetake = $true
            MaxRetakeAttempts = 3
            ComplianceRisk = "High"
            DataClassification = "Internal"
            DistributionScope = "All Employees"
            EffectiveDate = (Get-Date "2024-02-01")
            NextReviewDate = (Get-Date "2025-02-01")
            PolicyOwner = "David Nkosi"
            Department = "Operations"
            PolicySummary = "Workplace safety standards, emergency procedures, first aid protocols, and incident reporting requirements."
            EstimatedReadTimeMinutes = 12
        },
        @{
            Title = "Acceptable Use Policy"
            PolicyNumber = "POL-IT-002"
            PolicyName = "Acceptable Use Policy"
            PolicyCategory = "IT & Security"
            PolicyType = "Corporate"
            PolicyStatus = "Published"
            PolicyDescription = "Rules governing the acceptable use of company IT systems, networks, and devices."
            VersionNumber = "2.0"
            IsActive = $true
            IsMandatory = $true
            RequiresAcknowledgement = $true
            RequiresQuiz = $false
            ComplianceRisk = "Medium"
            DataClassification = "Internal"
            DistributionScope = "All Employees"
            EffectiveDate = (Get-Date "2024-04-15")
            NextReviewDate = (Get-Date "2025-04-15")
            PolicyOwner = "Sarah Johnson"
            Department = "Information Technology"
            PolicySummary = "Guidelines for appropriate use of company email, internet, devices, and software. Covers BYOD, remote work, and social media."
            EstimatedReadTimeMinutes = 8
        }
    )

    foreach ($policy in $testPolicies) {
        Write-Host "  Adding policy: $($policy.PolicyNumber) - $($policy.PolicyName)" -ForegroundColor Green
        Add-PnPListItem -List "PM_Policies" -Values $policy | Out-Null
    }

    Write-Host "  Test data inserted successfully." -ForegroundColor Green
}


# ============================================================================
# SUMMARY
# ============================================================================
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  Provisioning Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Lists created:" -ForegroundColor White
Write-Host "    1. PM_Policies              (List)     — Core policy records"
Write-Host "    2. PM_PolicyDocuments        (Library)  — Policy PDF/Word files"
Write-Host "    3. PM_PolicyAcknowledgements (List)     — User acknowledgements"
Write-Host "    4. PM_PolicyVersions         (List)     — Version history"
Write-Host "    5. PM_PolicyExemptions       (List)     — Exemption requests"
Write-Host "    6. PM_PolicySourceDocuments  (Library)  — Bulk import staging"
Write-Host "    7. PM_PolicyAuditLog         (List)     — Audit trail"
Write-Host "    8. PM_PolicyReadReceipts     (List)     — Read receipts"
Write-Host "    9. PM_PolicyQuizzes          (List)     — Quiz definitions"
Write-Host "   10. PM_PolicyQuizQuestions    (List)     — Quiz questions"
Write-Host "   11. PM_PolicyQuizResults      (List)     — Quiz results"
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor Yellow
Write-Host "    1. Upload policy PDFs to PM_PolicyDocuments"
Write-Host "    2. Set DocumentURL field on PM_Policies to link to documents"
Write-Host "       e.g., /sites/PolicyManager/PM_PolicyDocuments/YourPolicy.pdf"
Write-Host "    3. Deploy the .sppkg to the App Catalog"
Write-Host ""
