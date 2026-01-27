# ============================================================================
# DWx Policy Manager - Exemption & Distribution Lists
# Part 3: PM_PolicyExemptions, PM_PolicyDistributions, PM_PolicyTemplates
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
# LIST 7: PM_PolicyExemptions
# ============================================================================
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

# ============================================================================
# LIST 8: PM_PolicyDistributions
# ============================================================================
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

# ============================================================================
# LIST 9: PM_PolicyTemplates
# ============================================================================
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

Write-Host "`n✅ Exemption & Distribution lists created successfully!" -ForegroundColor Green
Write-Host "   - PM_PolicyExemptions" -ForegroundColor White
Write-Host "   - PM_PolicyDistributions" -ForegroundColor White
Write-Host "   - PM_PolicyTemplates" -ForegroundColor White
