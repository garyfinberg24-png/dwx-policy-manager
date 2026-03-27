# ============================================================================
# Policy Manager — Master Missing Lists Provisioning
# Creates all lists referenced in code but not yet provisioned.
# Assumes: Already connected via Connect-PnPOnline
# Safe to re-run — all commands use -ErrorAction SilentlyContinue
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Policy Manager — Missing Lists Master Provisioning" -ForegroundColor Cyan
Write-Host "  Creates lists referenced in code but not yet provisioned" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$created = 0
$skipped = 0

function Ensure-List {
    param([string]$Name, [string]$Description, [switch]$IsDocLib)
    $list = Get-PnPList -Identity $Name -ErrorAction SilentlyContinue
    if ($null -eq $list) {
        if ($IsDocLib) {
            New-PnPList -Title $Name -Template DocumentLibrary -EnableVersioning -ErrorAction SilentlyContinue | Out-Null
        } else {
            New-PnPList -Title $Name -Template GenericList -EnableVersioning -ErrorAction SilentlyContinue | Out-Null
        }
        Write-Host "  + Created: $Name" -ForegroundColor Green
        $script:created++
    } else {
        Write-Host "  ~ Exists:  $Name" -ForegroundColor Gray
        $script:skipped++
    }
}

# ============================================================================
# 1. PM_Configuration (used by Admin Centre, AI Chat, PolicyBuilder, BulkUpload)
# ============================================================================
Write-Host "`n[1/10] PM_Configuration" -ForegroundColor Yellow
Ensure-List -Name "PM_Configuration"
Add-PnPField -List "PM_Configuration" -DisplayName "Config Key" -InternalName "ConfigKey" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Configuration" -DisplayName "Config Value" -InternalName "ConfigValue" -Type Note -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Configuration" -DisplayName "Category" -InternalName "Category" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Configuration" -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_Configuration" -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List "PM_Configuration" -DisplayName "Is System Config" -InternalName "IsSystemConfig" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Configuration" -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_Configuration" -Identity "ConfigKey" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List "PM_Configuration" -Identity "Category" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "    Fields configured" -ForegroundColor Gray

# ============================================================================
# 2. PM_EmailQueue (legacy name — some services still reference it)
#    Code should use PM_NotificationQueue, but EmailQueueService still has
#    hardcoded PM_EmailQueue references. Create as alias/fallback.
# ============================================================================
Write-Host "`n[2/10] PM_EmailQueue (legacy fallback)" -ForegroundColor Yellow
Ensure-List -Name "PM_EmailQueue"
Add-PnPField -List "PM_EmailQueue" -DisplayName "Recipient Email" -InternalName "RecipientEmail" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_EmailQueue" -DisplayName "Recipient Name" -InternalName "RecipientName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_EmailQueue" -DisplayName "Sender Name" -InternalName "SenderName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_EmailQueue" -DisplayName "Sender Email" -InternalName "SenderEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_EmailQueue" -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_EmailQueue" -DisplayName "Policy Title" -InternalName "PolicyTitle" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_EmailQueue" -DisplayName "Notification Type" -InternalName "NotificationType" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_EmailQueue" -DisplayName "Channel" -InternalName "Channel" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_EmailQueue" -DisplayName "Message" -InternalName "Message" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_EmailQueue" -DisplayName "Queue Status" -InternalName "QueueStatus" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_EmailQueue" -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices "Low","Normal","High","Urgent" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_EmailQueue" -Identity "QueueStatus" -Values @{DefaultValue="Pending"} -ErrorAction SilentlyContinue
Set-PnPField -List "PM_EmailQueue" -Identity "QueueStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "    Fields configured" -ForegroundColor Gray

# ============================================================================
# 3. PM_HelpArticles (Help Centre — article content)
# ============================================================================
Write-Host "`n[3/10] PM_HelpArticles" -ForegroundColor Yellow
Ensure-List -Name "PM_HelpArticles"
Add-PnPField -List "PM_HelpArticles" -DisplayName "Article Body" -InternalName "ArticleBody" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpArticles" -DisplayName "Category" -InternalName "Category" -Type Choice -Choices "Getting Started","Policy Hub","My Policies","Author","Manager","Admin","Quiz","Distribution","Search","Troubleshooting" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpArticles" -DisplayName "Sort Order" -InternalName "SortOrder" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpArticles" -DisplayName "Is Featured" -InternalName "IsFeatured" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpArticles" -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_HelpArticles" -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List "PM_HelpArticles" -DisplayName "Tags" -InternalName "Tags" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpArticles" -DisplayName "View Count" -InternalName "ViewCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_HelpArticles" -Identity "Category" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "    Fields configured" -ForegroundColor Gray

# ============================================================================
# 4. PM_Cheatsheets (Help Centre — quick reference cards)
# ============================================================================
Write-Host "`n[4/10] PM_Cheatsheets" -ForegroundColor Yellow
Ensure-List -Name "PM_Cheatsheets"
Add-PnPField -List "PM_Cheatsheets" -DisplayName "Content" -InternalName "Content" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Cheatsheets" -DisplayName "Category" -InternalName "Category" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Cheatsheets" -DisplayName "Sort Order" -InternalName "SortOrder" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Cheatsheets" -DisplayName "Icon Name" -InternalName "IconName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Cheatsheets" -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_Cheatsheets" -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Write-Host "    Fields configured" -ForegroundColor Gray

# ============================================================================
# 5. PM_HelpTickets (Help Centre — support requests)
# ============================================================================
Write-Host "`n[5/10] PM_HelpTickets" -ForegroundColor Yellow
Ensure-List -Name "PM_HelpTickets"
Add-PnPField -List "PM_HelpTickets" -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpTickets" -DisplayName "Category" -InternalName "Category" -Type Choice -Choices "Bug","Feature Request","Question","Access Issue","Other" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpTickets" -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices "Low","Normal","High","Urgent" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpTickets" -DisplayName "Ticket Status" -InternalName "TicketStatus" -Type Choice -Choices "Open","In Progress","Resolved","Closed" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_HelpTickets" -Identity "TicketStatus" -Values @{DefaultValue="Open"} -ErrorAction SilentlyContinue
Add-PnPField -List "PM_HelpTickets" -DisplayName "Submitted By" -InternalName "SubmittedBy" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpTickets" -DisplayName "Submitted By Email" -InternalName "SubmittedByEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpTickets" -DisplayName "Assigned To" -InternalName "AssignedTo" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpTickets" -DisplayName "Resolution" -InternalName "Resolution" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_HelpTickets" -DisplayName "Resolved Date" -InternalName "ResolvedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_HelpTickets" -Identity "TicketStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "    Fields configured" -ForegroundColor Gray

# ============================================================================
# 6. PM_Delegations (fix — should match PM_ApprovalDelegations schema)
#    The constant PM_LISTS.DELEGATIONS pointed to 'PM_Delegations' but
#    code actually uses 'PM_ApprovalDelegations'. Create PM_Delegations
#    as a fallback with matching schema.
# ============================================================================
Write-Host "`n[6/10] PM_Delegations (fallback for workflow constant)" -ForegroundColor Yellow
Ensure-List -Name "PM_Delegations"
Add-PnPField -List "PM_Delegations" -DisplayName "Delegated By ID" -InternalName "DelegatedById" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Delegations" -DisplayName "Delegated By Name" -InternalName "DelegatedByName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Delegations" -DisplayName "Delegated To ID" -InternalName "DelegatedToId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Delegations" -DisplayName "Delegated To Name" -InternalName "DelegatedToName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Delegations" -DisplayName "Delegation Type" -InternalName "DelegationType" -Type Choice -Choices "Approval","Review","Full" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Delegations" -DisplayName "Start Date" -InternalName "StartDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Delegations" -DisplayName "End Date" -InternalName "EndDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Delegations" -DisplayName "Reason" -InternalName "Reason" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Delegations" -DisplayName "Delegation Status" -InternalName "DelegationStatus" -Type Choice -Choices "Active","Expired","Revoked" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_Delegations" -Identity "DelegationStatus" -Values @{DefaultValue="Active"} -ErrorAction SilentlyContinue
Add-PnPField -List "PM_Delegations" -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_Delegations" -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Set-PnPField -List "PM_Delegations" -Identity "DelegatedById" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List "PM_Delegations" -Identity "DelegatedToId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "    Fields configured" -ForegroundColor Gray

# ============================================================================
# 7. PM_WorkflowTemplates (Workflow engine — approval chain definitions)
# ============================================================================
Write-Host "`n[7/10] PM_WorkflowTemplates" -ForegroundColor Yellow
Ensure-List -Name "PM_WorkflowTemplates"
Add-PnPField -List "PM_WorkflowTemplates" -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowTemplates" -DisplayName "Workflow Type" -InternalName "WorkflowType" -Type Choice -Choices "Sequential","Parallel","Conditional" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowTemplates" -DisplayName "Steps JSON" -InternalName "StepsJSON" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowTemplates" -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_WorkflowTemplates" -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List "PM_WorkflowTemplates" -DisplayName "Is Default" -InternalName "IsDefault" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowTemplates" -DisplayName "SLA Days" -InternalName "SLADays" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowTemplates" -DisplayName "Escalation Action" -InternalName "EscalationAction" -Type Choice -Choices "Notify","Reassign","AutoApprove","Reject" -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields configured" -ForegroundColor Gray

# ============================================================================
# 8. PM_WorkflowInstances (Active workflow runs)
# ============================================================================
Write-Host "`n[8/10] PM_WorkflowInstances" -ForegroundColor Yellow
Ensure-List -Name "PM_WorkflowInstances"
Add-PnPField -List "PM_WorkflowInstances" -DisplayName "Template ID" -InternalName "TemplateId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowInstances" -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowInstances" -DisplayName "Policy Title" -InternalName "PolicyTitle" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowInstances" -DisplayName "Current Step" -InternalName "CurrentStep" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowInstances" -DisplayName "Workflow Status" -InternalName "WorkflowStatus" -Type Choice -Choices "Running","Completed","Cancelled","Failed","Escalated" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_WorkflowInstances" -Identity "WorkflowStatus" -Values @{DefaultValue="Running"} -ErrorAction SilentlyContinue
Add-PnPField -List "PM_WorkflowInstances" -DisplayName "Started By" -InternalName "StartedBy" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowInstances" -DisplayName "Started At" -InternalName "StartedAt" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowInstances" -DisplayName "Completed At" -InternalName "CompletedAt" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowInstances" -DisplayName "State JSON" -InternalName "StateJSON" -Type Note -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_WorkflowInstances" -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List "PM_WorkflowInstances" -Identity "WorkflowStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "    Fields configured" -ForegroundColor Gray

# ============================================================================
# 9. PM_WorkflowHistory (Audit trail for workflow actions)
# ============================================================================
Write-Host "`n[9/10] PM_WorkflowHistory" -ForegroundColor Yellow
Ensure-List -Name "PM_WorkflowHistory"
Add-PnPField -List "PM_WorkflowHistory" -DisplayName "Instance ID" -InternalName "InstanceId" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowHistory" -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowHistory" -DisplayName "Action" -InternalName "HistoryAction" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowHistory" -DisplayName "Action By" -InternalName "ActionBy" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowHistory" -DisplayName "Action Date" -InternalName "ActionDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowHistory" -DisplayName "Step Number" -InternalName "StepNumber" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowHistory" -DisplayName "Comments" -InternalName "Comments" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowHistory" -DisplayName "Previous Status" -InternalName "PreviousStatus" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_WorkflowHistory" -DisplayName "New Status" -InternalName "NewStatus" -Type Text -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List "PM_WorkflowHistory" -Identity "InstanceId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Write-Host "    Fields configured" -ForegroundColor Gray

# ============================================================================
# 10. Ensure PM_NotificationQueue has QueueStatus column
#     (fixes "Column 'Status' does not exist" error —
#      some services query 'Status' but the column was renamed to 'QueueStatus')
# ============================================================================
Write-Host "`n[10/10] PM_NotificationQueue — ensure QueueStatus column" -ForegroundColor Yellow
$nqList = Get-PnPList -Identity "PM_NotificationQueue" -ErrorAction SilentlyContinue
if ($null -ne $nqList) {
    Add-PnPField -List "PM_NotificationQueue" -DisplayName "Queue Status" -InternalName "QueueStatus" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
    Set-PnPField -List "PM_NotificationQueue" -Identity "QueueStatus" -Values @{DefaultValue="Pending"; Indexed=$true} -ErrorAction SilentlyContinue
    Write-Host "  QueueStatus column ensured on PM_NotificationQueue" -ForegroundColor Green
} else {
    Write-Host "  PM_NotificationQueue not found — run 07-Notification-Lists.ps1 first" -ForegroundColor Red
}

# ============================================================================
# ALSO: Ensure missing columns on existing lists
# ============================================================================
Write-Host "`n── Patching existing lists ──" -ForegroundColor Yellow

# PM_Notifications — ensure IsRead, Priority, ActionUrl columns exist
Write-Host "  Patching PM_Notifications..." -ForegroundColor Gray
Add-PnPField -List "PM_Notifications" -DisplayName "Is Read" -InternalName "IsRead" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Notifications" -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices "Low","Normal","High","Urgent" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Notifications" -DisplayName "Action URL" -InternalName "ActionUrl" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Notifications" -DisplayName "Related Item ID" -InternalName "RelatedItemId" -Type Number -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyAcknowledgements — ensure AckUserId column exists (was UserId in old schema)
Write-Host "  Patching PM_PolicyAcknowledgements..." -ForegroundColor Gray
Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "Ack User ID" -InternalName "AckUserId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "Policy Name" -InternalName "PolicyName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "Assigned Date" -InternalName "AssignedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_PolicyAcknowledgements" -DisplayName "Is Mandatory" -InternalName "IsMandatory" -Type Boolean -ErrorAction SilentlyContinue | Out-Null

# PM_Policies — ensure DocumentURL, DocumentFormat columns exist for Bulk Upload
Write-Host "  Patching PM_Policies..." -ForegroundColor Gray
Add-PnPField -List "PM_Policies" -DisplayName "Document URL" -InternalName "DocumentURL" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Policies" -DisplayName "Document Format" -InternalName "DocumentFormat" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Policies" -DisplayName "HTML Content" -InternalName "HTMLContent" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Policies" -DisplayName "Policy Content" -InternalName "PolicyContent" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Policies" -DisplayName "Policy Description" -InternalName "PolicyDescription" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Policies" -DisplayName "Read Timeframe" -InternalName "ReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6","Custom" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Policies" -DisplayName "Next Review Date" -InternalName "NextReviewDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Policies" -DisplayName "Review Frequency" -InternalName "ReviewFrequency" -Type Choice -Choices "Monthly","Quarterly","Biannual","Annual","None" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Policies" -DisplayName "Requires Acknowledgement" -InternalName "RequiresAcknowledgement" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List "PM_Policies" -DisplayName "Requires Quiz" -InternalName "RequiresQuiz" -Type Boolean -ErrorAction SilentlyContinue | Out-Null

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Missing Lists Provisioning Complete" -ForegroundColor Green
Write-Host "  Created: $created | Already existed: $skipped" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Lists created:" -ForegroundColor White
Write-Host "    1. PM_Configuration        — Admin config key-value store" -ForegroundColor White
Write-Host "    2. PM_EmailQueue            — Legacy email queue fallback" -ForegroundColor White
Write-Host "    3. PM_HelpArticles          — Help Centre articles" -ForegroundColor White
Write-Host "    4. PM_Cheatsheets           — Help Centre quick references" -ForegroundColor White
Write-Host "    5. PM_HelpTickets           — Help Centre support tickets" -ForegroundColor White
Write-Host "    6. PM_Delegations           — Workflow delegations" -ForegroundColor White
Write-Host "    7. PM_WorkflowTemplates     — Approval chain templates" -ForegroundColor White
Write-Host "    8. PM_WorkflowInstances     — Active workflow runs" -ForegroundColor White
Write-Host "    9. PM_WorkflowHistory       — Workflow audit trail" -ForegroundColor White
Write-Host "   10. PM_NotificationQueue     — QueueStatus column patch" -ForegroundColor White
Write-Host ""
Write-Host "  Columns patched on:" -ForegroundColor White
Write-Host "    - PM_Notifications          — IsRead, Priority, ActionUrl" -ForegroundColor White
Write-Host "    - PM_PolicyAcknowledgements  — AckUserId, PolicyName, AssignedDate" -ForegroundColor White
Write-Host "    - PM_Policies               — DocumentURL, DocumentFormat, content fields" -ForegroundColor White
Write-Host ""
