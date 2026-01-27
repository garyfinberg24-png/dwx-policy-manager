# ============================================================================
# DWx Policy Manager - Analytics & Audit Lists
# Part 6: PM_PolicyAuditLog, PM_PolicyAnalytics, PM_PolicyFeedback, PM_PolicyDocuments
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

# ============================================================================
# LIST 17: PM_PolicyAuditLog
# ============================================================================
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

# ============================================================================
# LIST 18: PM_PolicyAnalytics
# ============================================================================
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

# ============================================================================
# LIST 19: PM_PolicyFeedback
# ============================================================================
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

# ============================================================================
# LIST 20: PM_PolicyDocuments
# ============================================================================
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

Write-Host "`n Analytics & Audit lists created successfully!" -ForegroundColor Green
Write-Host "   - PM_PolicyAuditLog" -ForegroundColor White
Write-Host "   - PM_PolicyAnalytics" -ForegroundColor White
Write-Host "   - PM_PolicyFeedback" -ForegroundColor White
Write-Host "   - PM_PolicyDocuments" -ForegroundColor White
