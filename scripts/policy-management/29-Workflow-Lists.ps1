# ============================================================================
# Policy Manager — Workflow Engine Lists
# Provisions PM_WorkflowTemplates, PM_WorkflowInstances, PM_EscalationRules
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Workflow Engine Lists Provisioning" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ============================================================================
# PM_WorkflowTemplates — Reusable approval workflow templates
# ============================================================================

$listName = "PM_WorkflowTemplates"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Template Name" -InternalName "TemplateName" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Workflow Type" -InternalName "WorkflowType" -Type Choice -Choices "FastTrack","Standard","Regulatory","Custom" -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Approval Levels" -InternalName "ApprovalLevels" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Level Definitions" -InternalName "LevelDefinitions" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Escalation Enabled" -InternalName "EscalationEnabled" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Escalation Days" -InternalName "EscalationDays" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Default" -InternalName "IsDefault" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Created By Email" -InternalName "CreatedByEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Created Date" -InternalName "TemplateCreatedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null

# Index on WorkflowType for filtered queries
Set-PnPField -List $listName -Identity "WorkflowType" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  Fields added to: $listName" -ForegroundColor Green
Write-Host ""

# ============================================================================
# PM_WorkflowInstances — Running approval workflow instances
# ============================================================================

$listName = "PM_WorkflowInstances"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Policy Title" -InternalName "PolicyTitle" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Template ID" -InternalName "TemplateId" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Template Name" -InternalName "TemplateName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Current Level" -InternalName "CurrentLevel" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Total Levels" -InternalName "TotalLevels" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Workflow Status" -InternalName "WorkflowStatus" -Type Choice -Choices "Active","Completed","Cancelled","Escalated" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Started Date" -InternalName "StartedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Completed Date" -InternalName "CompletedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Started By" -InternalName "StartedBy" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Escalation Count" -InternalName "EscalationCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "EscalationCount" -Values @{DefaultValue="0"} -ErrorAction SilentlyContinue

# Indexes
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "WorkflowStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  Fields added to: $listName" -ForegroundColor Green
Write-Host ""

# ============================================================================
# PM_EscalationRules — Escalation rule definitions per template/level
# ============================================================================

$listName = "PM_EscalationRules"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Template ID" -InternalName "TemplateId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Level" -InternalName "Level" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Timeout Days" -InternalName "TimeoutDays" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Escalation Action" -InternalName "EscalationAction" -Type Choice -Choices "Reassign","NotifyManager","AutoApprove","Reject" -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Escalate To Email" -InternalName "EscalateToEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue

# Index on TemplateId for filtered queries
Set-PnPField -List $listName -Identity "TemplateId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  Fields added to: $listName" -ForegroundColor Green
Write-Host ""

# ============================================================================
# Add WorkflowTemplateId column to PM_Policies (optional field)
# ============================================================================

$policiesList = "PM_Policies"
$existingField = Get-PnPField -List $policiesList -Identity "WorkflowTemplateId" -ErrorAction SilentlyContinue

if ($null -eq $existingField) {
    Add-PnPField -List $policiesList -DisplayName "Workflow Template ID" -InternalName "WorkflowTemplateId" -Type Number -ErrorAction SilentlyContinue | Out-Null
    Write-Host "  Added WorkflowTemplateId column to $policiesList" -ForegroundColor Green
} else {
    Write-Host "  WorkflowTemplateId column already exists on $policiesList" -ForegroundColor Gray
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Workflow Engine Lists — Complete" -ForegroundColor Green
Write-Host "  Lists: PM_WorkflowTemplates, PM_WorkflowInstances, PM_EscalationRules" -ForegroundColor Green
Write-Host "  Column: WorkflowTemplateId added to PM_Policies" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
