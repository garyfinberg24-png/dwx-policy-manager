# ============================================================================
# DWx Policy Manager - Approval Lists
# Part 8: PM_Approvals, PM_ApprovalChains, PM_ApprovalHistory,
#          PM_ApprovalDelegations, PM_ApprovalTemplates
#
# These are the lists used by ApprovalService.ts in the SPFx solution.
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
# Helper: Create list if it doesn't exist
# ============================================================================
function Ensure-List {
    param(
        [string]$ListName,
        [string]$Description = ""
    )
    $list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($null -eq $list) {
        New-PnPList -Title $ListName -Template GenericList -EnableVersioning
        Write-Host "  Created list: $ListName" -ForegroundColor Green
    } else {
        Write-Host "  List already exists: $ListName" -ForegroundColor Gray
    }
}

# ============================================================================
# LIST: PM_Approvals
# Individual approval records (one per approver per level)
# Used by ApprovalService.ts -> this.APPROVALS_LIST
# ============================================================================
Write-Host "`n Creating PM_Approvals list..." -ForegroundColor Yellow

$listName = "PM_Approvals"
Ensure-List -ListName $listName

Add-PnPField -List $listName -DisplayName "Process ID" -InternalName "ProcessID" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approval Chain ID" -InternalName "ApprovalChainId" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approval Level" -InternalName "ApprovalLevel" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approval Sequence" -InternalName "ApprovalSequence" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approval Type" -InternalName "ApprovalType" -Type Choice -Choices "Sequential","Parallel","FirstApprover" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Pending","Approved","Rejected","Delegated","Escalated","Cancelled","Skipped","Queued","Expired" -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approver ID" -InternalName "ApproverId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Original Approver ID" -InternalName "OriginalApproverId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Delegated By ID" -InternalName "DelegatedById" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Requested Date" -InternalName "RequestedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Due Date" -InternalName "DueDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Completed Date" -InternalName "CompletedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Response Time" -InternalName "ResponseTime" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Decision" -InternalName "Decision" -Type Choice -Choices "Approved","Rejected" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Comments" -InternalName "Comments" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Notes" -InternalName "Notes" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Reason Required" -InternalName "ReasonRequired" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Overdue" -InternalName "IsOverdue" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Escalation Level" -InternalName "EscalationLevel" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Escalation Date" -InternalName "EscalationDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Escalation Action" -InternalName "EscalationAction" -Type Choice -Choices "Notify","AutoApprove","AssignToManager","AssignToAlternate" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Workflow Instance ID" -InternalName "WorkflowInstanceId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Workflow Step ID" -InternalName "WorkflowStepId" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approval Template ID" -InternalName "ApprovalTemplateId" -Type Number -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "ProcessID" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "ApproverId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "Status" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "DueDate" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_Approvals list configured" -ForegroundColor Green

# ============================================================================
# LIST: PM_ApprovalChains
# Approval chain definitions (one per process approval)
# Used by ApprovalService.ts -> this.APPROVAL_CHAINS_LIST
# ============================================================================
Write-Host "`n Creating PM_ApprovalChains list..." -ForegroundColor Yellow

$listName = "PM_ApprovalChains"
Ensure-List -ListName $listName

Add-PnPField -List $listName -DisplayName "Process ID" -InternalName "ProcessID" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Chain Name" -InternalName "ChainName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approval Type" -InternalName "ApprovalType" -Type Choice -Choices "Sequential","Parallel","FirstApprover" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Levels" -InternalName "Levels" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Require Comments" -InternalName "RequireComments" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Allow Delegation" -InternalName "AllowDelegation" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Auto Escalation Days" -InternalName "AutoEscalationDays" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Escalation Action" -InternalName "EscalationAction" -Type Choice -Choices "Notify","AutoApprove","AssignToManager","AssignToAlternate" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Current Level" -InternalName "CurrentLevel" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Overall Status" -InternalName "OverallStatus" -Type Choice -Choices "Pending","Approved","Rejected","Delegated","Escalated","Cancelled","Skipped","Queued","Expired" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Start Date" -InternalName "StartDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Completed Date" -InternalName "CompletedDate" -Type DateTime -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "ProcessID" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "OverallStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_ApprovalChains list configured" -ForegroundColor Green

# ============================================================================
# LIST: PM_ApprovalHistory
# Audit trail of approval actions
# Used by ApprovalService.ts -> this.APPROVAL_HISTORY_LIST
# ============================================================================
Write-Host "`n Creating PM_ApprovalHistory list..." -ForegroundColor Yellow

$listName = "PM_ApprovalHistory"
Ensure-List -ListName $listName

Add-PnPField -List $listName -DisplayName "Approval ID" -InternalName "ApprovalId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Process ID" -InternalName "ProcessID" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Action" -InternalName "Action" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Performed By ID" -InternalName "PerformedById" -Type Number -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Action Date" -InternalName "ActionDate" -Type DateTime -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Comments" -InternalName "Comments" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Notes" -InternalName "Notes" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Previous Status" -InternalName "PreviousStatus" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "New Status" -InternalName "NewStatus" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Delegated To ID" -InternalName "DelegatedToId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Delegation Reason" -InternalName "DelegationReason" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Escalated To ID" -InternalName "EscalatedToId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Escalation Reason" -InternalName "EscalationReason" -Type Note -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "ApprovalId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "ProcessID" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PerformedById" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_ApprovalHistory list configured" -ForegroundColor Green

# ============================================================================
# LIST: PM_ApprovalDelegations
# Delegation assignments
# Used by ApprovalService.ts -> this.DELEGATIONS_LIST
# ============================================================================
Write-Host "`n Creating PM_ApprovalDelegations list..." -ForegroundColor Yellow

$listName = "PM_ApprovalDelegations"
Ensure-List -ListName $listName

Add-PnPField -List $listName -DisplayName "Delegated By ID" -InternalName "DelegatedById" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Delegated To ID" -InternalName "DelegatedToId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Start Date" -InternalName "StartDate" -Type DateTime -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "End Date" -InternalName "EndDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Reason" -InternalName "Reason" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Process Types" -InternalName "ProcessTypes" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Auto Delegate" -InternalName "AutoDelegate" -Type Boolean -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "DelegatedById" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "DelegatedToId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_ApprovalDelegations list configured" -ForegroundColor Green

# ============================================================================
# LIST: PM_ApprovalTemplates
# Reusable approval workflow templates
# Used by ApprovalService.ts -> this.TEMPLATES_LIST
# ============================================================================
Write-Host "`n Creating PM_ApprovalTemplates list..." -ForegroundColor Yellow

$listName = "PM_ApprovalTemplates"
Ensure-List -ListName $listName

Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Process Types" -InternalName "ProcessTypes" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Approval Type" -InternalName "ApprovalType" -Type Choice -Choices "Sequential","Parallel","FirstApprover" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Levels" -InternalName "Levels" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Require Comments" -InternalName "RequireComments" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Allow Delegation" -InternalName "AllowDelegation" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Auto Escalation Days" -InternalName "AutoEscalationDays" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Escalation Action" -InternalName "EscalationAction" -Type Choice -Choices "Notify","AutoApprove","AssignToManager","AssignToAlternate" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_ApprovalTemplates list configured" -ForegroundColor Green

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n Approval lists created successfully!" -ForegroundColor Green
Write-Host "   - PM_Approvals            (individual approval records)" -ForegroundColor White
Write-Host "   - PM_ApprovalChains       (approval chain instances)" -ForegroundColor White
Write-Host "   - PM_ApprovalHistory      (action audit trail)" -ForegroundColor White
Write-Host "   - PM_ApprovalDelegations  (delegation assignments)" -ForegroundColor White
Write-Host "   - PM_ApprovalTemplates    (reusable workflow templates)" -ForegroundColor White

Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "  1. Run Seed-ApprovalAndNotificationData.ps1 to populate sample data" -ForegroundColor Gray
Write-Host "  2. Deploy the updated SPFx package" -ForegroundColor Gray
