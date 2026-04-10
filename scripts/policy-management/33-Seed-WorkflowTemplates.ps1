# ============================================================================
# 33-Seed-WorkflowTemplates.ps1
# Seeds 6 reusable multi-level approval workflow templates
# ============================================================================
# PREREQUISITE: Already connected to SharePoint via Connect-PnPOnline
# ============================================================================

$listName = "PM_WorkflowTemplates"

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Seeding Workflow Templates — $listName" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Cyan

# Check list exists
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "  ✗ $listName does not exist — creating..." -ForegroundColor Yellow
    New-PnPList -Title $listName -Template GenericList -EnableVersioning -OnQuickLaunch:$false
    # Add required columns
    Add-PnPField -List $listName -DisplayName "TemplateName" -InternalName "TemplateName" -Type Text
    Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note
    Add-PnPField -List $listName -DisplayName "WorkflowType" -InternalName "WorkflowType" -Type Choice -Choices "Sequential","Parallel","Hybrid"
    Add-PnPField -List $listName -DisplayName "Levels" -InternalName "Levels" -Type Note
    Add-PnPField -List $listName -DisplayName "ApplicableTo" -InternalName "ApplicableTo" -Type Text
    Add-PnPField -List $listName -DisplayName "SLADays" -InternalName "SLADays" -Type Number
    Add-PnPField -List $listName -DisplayName "EscalationAction" -InternalName "EscalationAction" -Type Choice -Choices "Notify","Reassign","AutoApprove","Reject"
    Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean
    Add-PnPField -List $listName -DisplayName "IsDefault" -InternalName "IsDefault" -Type Boolean
    Add-PnPField -List $listName -DisplayName "RequiresAllApprovers" -InternalName "RequiresAllApprovers" -Type Boolean
    Write-Host "  ✓ $listName created with columns" -ForegroundColor Green
}

$templates = @(
    @{
        Title = "Fast Track — Single Approver"
        TemplateName = "Fast Track — Single Approver"
        Description = "Quick approval for low-risk, operational policies. One approver reviews and approves. Ideal for minor updates, internal guidelines, and informational policies."
        WorkflowType = "Sequential"
        ApplicableTo = "Low Risk, Informational"
        SLADays = 3
        EscalationAction = "AutoApprove"
        IsActive = $true
        IsDefault = $true
        RequiresAllApprovers = $false
        Levels = '[{"level":1,"name":"Department Approver","approverType":"Role","approverValue":"Manager","required":true,"sla":3}]'
    },
    @{
        Title = "Standard — Two-Level Review"
        TemplateName = "Standard — Two-Level Review"
        Description = "Standard two-level workflow for most policies. First reviewed by a subject matter expert, then approved by a department manager. Covers medium-risk HR, IT, and operational policies."
        WorkflowType = "Sequential"
        ApplicableTo = "Medium Risk, HR, IT, Operational"
        SLADays = 7
        EscalationAction = "Notify"
        IsActive = $true
        IsDefault = $false
        RequiresAllApprovers = $true
        Levels = '[{"level":1,"name":"Subject Matter Expert Review","approverType":"Role","approverValue":"Author","required":true,"sla":3},{"level":2,"name":"Department Manager Approval","approverType":"Role","approverValue":"Manager","required":true,"sla":4}]'
    },
    @{
        Title = "Regulatory — Three-Level with Legal"
        TemplateName = "Regulatory — Three-Level with Legal"
        Description = "Rigorous three-level workflow for regulatory and compliance policies (POPIA, GDPR, SOX). Includes legal review as a mandatory step before final executive sign-off."
        WorkflowType = "Sequential"
        ApplicableTo = "Critical, High Risk, Regulatory, Compliance"
        SLADays = 14
        EscalationAction = "Notify"
        IsActive = $true
        IsDefault = $false
        RequiresAllApprovers = $true
        Levels = '[{"level":1,"name":"Compliance Officer Review","approverType":"Group","approverValue":"PM_ComplianceOfficers","required":true,"sla":5},{"level":2,"name":"Legal Review","approverType":"Group","approverValue":"PM_LegalReviewers","required":true,"sla":5},{"level":3,"name":"Executive Sign-Off","approverType":"Role","approverValue":"Admin","required":true,"sla":4}]'
    },
    @{
        Title = "Parallel — Committee Review"
        TemplateName = "Parallel — Committee Review"
        Description = "All committee members review simultaneously. Requires majority approval (any 3 of the assigned reviewers). Used for cross-departmental policies that need broad stakeholder input."
        WorkflowType = "Parallel"
        ApplicableTo = "Cross-Department, Strategic"
        SLADays = 10
        EscalationAction = "Notify"
        IsActive = $true
        IsDefault = $false
        RequiresAllApprovers = $false
        Levels = '[{"level":1,"name":"Committee Review","approverType":"Group","approverValue":"PM_PolicyCommittee","required":false,"minimumApprovals":3,"sla":10}]'
    },
    @{
        Title = "Executive — CEO/Board Approval"
        TemplateName = "Executive — CEO/Board Approval"
        Description = "High-stakes workflow for policies requiring C-suite or board approval. Includes a preliminary management review before escalation to executive leadership. Used for strategic, financial, and governance policies."
        WorkflowType = "Sequential"
        ApplicableTo = "Critical, Financial, Governance, Strategic"
        SLADays = 21
        EscalationAction = "Reassign"
        IsActive = $true
        IsDefault = $false
        RequiresAllApprovers = $true
        Levels = '[{"level":1,"name":"Department Head Review","approverType":"Role","approverValue":"Manager","required":true,"sla":5},{"level":2,"name":"Risk & Compliance Review","approverType":"Group","approverValue":"PM_ComplianceOfficers","required":true,"sla":5},{"level":3,"name":"CFO/COO Approval","approverType":"Group","approverValue":"PM_ExecutiveApprovers","required":true,"sla":7},{"level":4,"name":"CEO/Board Sign-Off","approverType":"Group","approverValue":"PM_BoardApprovers","required":true,"sla":4}]'
    },
    @{
        Title = "Emergency — Expedited with Auto-Escalation"
        TemplateName = "Emergency — Expedited with Auto-Escalation"
        Description = "Fast-track workflow for urgent policy changes (security incidents, regulatory deadlines). Single approver with 24-hour SLA and automatic escalation to Admin if not actioned. Used when immediate policy changes are required."
        WorkflowType = "Sequential"
        ApplicableTo = "Emergency, Security, Urgent"
        SLADays = 1
        EscalationAction = "Reassign"
        IsActive = $true
        IsDefault = $false
        RequiresAllApprovers = $false
        Levels = '[{"level":1,"name":"Immediate Manager Approval","approverType":"Role","approverValue":"Manager","required":true,"sla":1},{"level":2,"name":"Auto-Escalation to Admin","approverType":"Role","approverValue":"Admin","required":true,"sla":1,"isEscalation":true}]'
    }
)

foreach ($template in $templates) {
    # Check if template already exists
    $existing = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($template.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  - $($template.Title) already exists — skipping" -ForegroundColor DarkGray
        continue
    }

    try {
        Add-PnPListItem -List $listName -Values $template | Out-Null
        Write-Host "  ✓ $($template.Title)" -ForegroundColor Green
    } catch {
        Write-Host "  ✗ $($template.Title): $_" -ForegroundColor Red
    }
}

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  6 workflow templates seeded!" -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Green
