# ============================================================================
# Seed-PolicyApprovalSampleData.ps1
# Seeds Policy Approval Workflow lists with realistic sample data
#
# PREREQUISITE: You must already be connected to SharePoint via PnP PowerShell
# PREREQUISITE: Run Provision-PolicyApprovalLists.ps1 first
# ============================================================================

param(
    [switch]$WhatIf = $false
)

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Policy Approval Workflow - Sample Data Seeding" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

if ($WhatIf) {
    Write-Host "[WhatIf Mode] No changes will be made`n" -ForegroundColor Yellow
}

# ============================================================================
# SAMPLE USERS (Update these IDs to match your environment)
# ============================================================================

# Get some users from the site
Write-Host "Fetching site users..." -ForegroundColor White
$siteUsers = Get-PnPUser | Where-Object { $_.Email -ne "" } | Select-Object -First 10

if ($siteUsers.Count -lt 3) {
    Write-Host "Warning: Need at least 3 users with emails for sample data. Found: $($siteUsers.Count)" -ForegroundColor Yellow
    Write-Host "Sample data will use placeholder IDs" -ForegroundColor Yellow

    # Create placeholder user data
    $users = @(
        @{ Id = 1; Name = "Sarah Mitchell"; Email = "sarah.mitchell@contoso.com"; Title = "HR Director" },
        @{ Id = 2; Name = "James Chen"; Email = "james.chen@contoso.com"; Title = "Policy Manager" },
        @{ Id = 3; Name = "Emily Rodriguez"; Email = "emily.rodriguez@contoso.com"; Title = "Compliance Officer" },
        @{ Id = 4; Name = "Michael Thompson"; Email = "michael.thompson@contoso.com"; Title = "Legal Counsel" },
        @{ Id = 5; Name = "Amanda Foster"; Email = "amanda.foster@contoso.com"; Title = "IT Security Lead" },
        @{ Id = 6; Name = "David Kim"; Email = "david.kim@contoso.com"; Title = "Operations Manager" },
        @{ Id = 7; Name = "Jennifer Walsh"; Email = "jennifer.walsh@contoso.com"; Title = "Finance Director" },
        @{ Id = 8; Name = "Robert Garcia"; Email = "robert.garcia@contoso.com"; Title = "CEO" }
    )
} else {
    $users = $siteUsers | ForEach-Object {
        @{ Id = $_.Id; Name = $_.Title; Email = $_.Email; Title = "Team Member" }
    }
}

Write-Host "Using $($users.Count) users for sample data`n" -ForegroundColor Gray

# ============================================================================
# SEED: JML_Policy_ApprovalTemplates
# ============================================================================

$listName = "JML_Policy_ApprovalTemplates"
Write-Host "Seeding: $listName..." -ForegroundColor White

$approvalTemplates = @(
    @{
        Title = "Standard Policy Approval"
        TemplateName = "Standard Policy Approval"
        Description = "Standard two-stage approval workflow for general policies. First reviewed by department head, then approved by policy committee."
        Category = "General"
        IsDefault = $true
        IsActive = $true
        RequireAllStages = $true
        AllowParallelApproval = $false
        NotifyOnComplete = $true
        AutoArchiveOnComplete = $false
        StagesJson = '[
            {
                "stageNumber": 1,
                "stageName": "Department Review",
                "stageType": "Individual",
                "approverIds": [' + $users[1].Id + '],
                "approverNames": ["' + $users[1].Name + '"],
                "daysToComplete": 5,
                "isRequired": true,
                "canDelegate": true,
                "autoApproveOnTimeout": false
            },
            {
                "stageNumber": 2,
                "stageName": "Policy Committee Approval",
                "stageType": "AnyOf",
                "approverIds": [' + $users[0].Id + ', ' + $users[2].Id + '],
                "approverNames": ["' + $users[0].Name + '", "' + $users[2].Name + '"],
                "daysToComplete": 7,
                "minimumApprovers": 1,
                "isRequired": true,
                "canDelegate": true,
                "autoApproveOnTimeout": false
            }
        ]'
    },
    @{
        Title = "HR Policy - Three Stage"
        TemplateName = "HR Policy - Three Stage"
        Description = "Comprehensive three-stage approval for HR policies. Requires HR review, legal review, and executive sign-off."
        Category = "HR"
        IsDefault = $false
        IsActive = $true
        RequireAllStages = $true
        AllowParallelApproval = $false
        NotifyOnComplete = $true
        AutoArchiveOnComplete = $false
        StagesJson = '[
            {
                "stageNumber": 1,
                "stageName": "HR Review",
                "stageType": "Individual",
                "approverIds": [' + $users[0].Id + '],
                "approverNames": ["' + $users[0].Name + '"],
                "daysToComplete": 5,
                "isRequired": true,
                "canDelegate": true
            },
            {
                "stageNumber": 2,
                "stageName": "Legal Review",
                "stageType": "Individual",
                "approverIds": [' + $users[3].Id + '],
                "approverNames": ["' + $users[3].Name + '"],
                "daysToComplete": 7,
                "isRequired": true,
                "canDelegate": false
            },
            {
                "stageNumber": 3,
                "stageName": "Executive Approval",
                "stageType": "Individual",
                "approverIds": [' + $users[7].Id + '],
                "approverNames": ["' + $users[7].Name + '"],
                "daysToComplete": 3,
                "isRequired": true,
                "canDelegate": true
            }
        ]'
    },
    @{
        Title = "IT Security Policy"
        TemplateName = "IT Security Policy"
        Description = "Two-stage parallel approval for IT security policies. IT Security and Compliance must both approve."
        Category = "IT"
        IsDefault = $false
        IsActive = $true
        RequireAllStages = $true
        AllowParallelApproval = $true
        NotifyOnComplete = $true
        AutoArchiveOnComplete = $false
        StagesJson = '[
            {
                "stageNumber": 1,
                "stageName": "Technical Review",
                "stageType": "AllOf",
                "approverIds": [' + $users[4].Id + ', ' + $users[2].Id + '],
                "approverNames": ["' + $users[4].Name + '", "' + $users[2].Name + '"],
                "daysToComplete": 10,
                "minimumApprovers": 2,
                "isRequired": true,
                "canDelegate": true
            },
            {
                "stageNumber": 2,
                "stageName": "Management Sign-off",
                "stageType": "Individual",
                "approverIds": [' + $users[5].Id + '],
                "approverNames": ["' + $users[5].Name + '"],
                "daysToComplete": 5,
                "isRequired": true,
                "canDelegate": true
            }
        ]'
    },
    @{
        Title = "Finance Policy - Dual Control"
        TemplateName = "Finance Policy - Dual Control"
        Description = "Finance policies requiring dual control approval. Two independent approvers from finance and compliance."
        Category = "Finance"
        IsDefault = $false
        IsActive = $true
        RequireAllStages = $true
        AllowParallelApproval = $true
        NotifyOnComplete = $true
        AutoArchiveOnComplete = $true
        StagesJson = '[
            {
                "stageNumber": 1,
                "stageName": "Finance Director Review",
                "stageType": "Individual",
                "approverIds": [' + $users[6].Id + '],
                "approverNames": ["' + $users[6].Name + '"],
                "daysToComplete": 5,
                "isRequired": true,
                "canDelegate": false
            },
            {
                "stageNumber": 2,
                "stageName": "Compliance Verification",
                "stageType": "Individual",
                "approverIds": [' + $users[2].Id + '],
                "approverNames": ["' + $users[2].Name + '"],
                "daysToComplete": 5,
                "isRequired": true,
                "canDelegate": false
            }
        ]'
    },
    @{
        Title = "Quick Approval - Single Stage"
        TemplateName = "Quick Approval - Single Stage"
        Description = "Expedited single-stage approval for low-risk policy updates. Any policy committee member can approve."
        Category = "General"
        IsDefault = $false
        IsActive = $true
        RequireAllStages = $true
        AllowParallelApproval = $false
        NotifyOnComplete = $true
        AutoArchiveOnComplete = $false
        StagesJson = '[
            {
                "stageNumber": 1,
                "stageName": "Policy Committee",
                "stageType": "AnyOf",
                "approverIds": [' + $users[0].Id + ', ' + $users[1].Id + ', ' + $users[2].Id + '],
                "approverNames": ["' + $users[0].Name + '", "' + $users[1].Name + '", "' + $users[2].Name + '"],
                "daysToComplete": 3,
                "minimumApprovers": 1,
                "isRequired": true,
                "canDelegate": true,
                "autoApproveOnTimeout": false
            }
        ]'
    },
    @{
        Title = "Legal & Compliance Full Review"
        TemplateName = "Legal & Compliance Full Review"
        Description = "Comprehensive four-stage approval for high-risk legal and compliance policies. Requires legal, compliance, HR, and executive approval."
        Category = "Legal"
        IsDefault = $false
        IsActive = $true
        RequireAllStages = $true
        AllowParallelApproval = $false
        NotifyOnComplete = $true
        AutoArchiveOnComplete = $true
        StagesJson = '[
            {
                "stageNumber": 1,
                "stageName": "Legal Draft Review",
                "stageType": "Individual",
                "approverIds": [' + $users[3].Id + '],
                "approverNames": ["' + $users[3].Name + '"],
                "daysToComplete": 7,
                "isRequired": true,
                "canDelegate": false
            },
            {
                "stageNumber": 2,
                "stageName": "Compliance Assessment",
                "stageType": "Individual",
                "approverIds": [' + $users[2].Id + '],
                "approverNames": ["' + $users[2].Name + '"],
                "daysToComplete": 5,
                "isRequired": true,
                "canDelegate": true
            },
            {
                "stageNumber": 3,
                "stageName": "HR Impact Review",
                "stageType": "Individual",
                "approverIds": [' + $users[0].Id + '],
                "approverNames": ["' + $users[0].Name + '"],
                "daysToComplete": 3,
                "isRequired": true,
                "canDelegate": true
            },
            {
                "stageNumber": 4,
                "stageName": "Executive Sign-off",
                "stageType": "Individual",
                "approverIds": [' + $users[7].Id + '],
                "approverNames": ["' + $users[7].Name + '"],
                "daysToComplete": 3,
                "isRequired": true,
                "canDelegate": true
            }
        ]'
    }
)

$templateCount = 0
foreach ($template in $approvalTemplates) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $template | Out-Null
            $templateCount++
            Write-Host "  Added: $($template.TemplateName)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($template.TemplateName): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($template.TemplateName)" -ForegroundColor Gray
    }
}
Write-Host "  Created $templateCount templates`n" -ForegroundColor Green

# ============================================================================
# SEED: JML_Policy_ApprovalWorkflows
# ============================================================================

$listName = "JML_Policy_ApprovalWorkflows"
Write-Host "Seeding: $listName..." -ForegroundColor White

$now = Get-Date
$workflows = @(
    @{
        Title = "WF-2024-001: Remote Work Policy v2.0"
        PolicyId = 101
        PolicyTitle = "Remote Work Policy v2.0"
        TemplateId = 1
        Status = "InProgress"
        CurrentStage = 2
        TotalStages = 2
        InitiatedById = $users[1].Id
        InitiatedByName = $users[1].Name
        InitiatedDate = $now.AddDays(-5).ToString("o")
        DueDate = $now.AddDays(9).ToString("o")
        IsUrgent = $false
        Priority = "Normal"
        StagesJson = '[
            {"stageNumber": 1, "stageName": "Department Review", "status": "Approved", "completedDate": "' + $now.AddDays(-3).ToString("o") + '"},
            {"stageNumber": 2, "stageName": "Policy Committee Approval", "status": "Pending", "dueDate": "' + $now.AddDays(4).ToString("o") + '"}
        ]'
    },
    @{
        Title = "WF-2024-002: Data Protection Policy"
        PolicyId = 102
        PolicyTitle = "Data Protection Policy"
        TemplateId = 3
        Status = "InProgress"
        CurrentStage = 1
        TotalStages = 2
        InitiatedById = $users[4].Id
        InitiatedByName = $users[4].Name
        InitiatedDate = $now.AddDays(-2).ToString("o")
        DueDate = $now.AddDays(13).ToString("o")
        IsUrgent = $true
        Priority = "High"
        StagesJson = '[
            {"stageNumber": 1, "stageName": "Technical Review", "status": "Pending", "dueDate": "' + $now.AddDays(8).ToString("o") + '"},
            {"stageNumber": 2, "stageName": "Management Sign-off", "status": "NotStarted"}
        ]'
    },
    @{
        Title = "WF-2024-003: Employee Benefits Policy"
        PolicyId = 103
        PolicyTitle = "Employee Benefits Policy"
        TemplateId = 2
        Status = "Approved"
        CurrentStage = 3
        TotalStages = 3
        InitiatedById = $users[0].Id
        InitiatedByName = $users[0].Name
        InitiatedDate = $now.AddDays(-21).ToString("o")
        DueDate = $now.AddDays(-6).ToString("o")
        CompletedDate = $now.AddDays(-8).ToString("o")
        FinalDecision = "Approved"
        FinalComments = "All stages approved. Policy ready for publication."
        IsUrgent = $false
        Priority = "Normal"
        StagesJson = '[
            {"stageNumber": 1, "stageName": "HR Review", "status": "Approved", "completedDate": "' + $now.AddDays(-18).ToString("o") + '"},
            {"stageNumber": 2, "stageName": "Legal Review", "status": "Approved", "completedDate": "' + $now.AddDays(-12).ToString("o") + '"},
            {"stageNumber": 3, "stageName": "Executive Approval", "status": "Approved", "completedDate": "' + $now.AddDays(-8).ToString("o") + '"}
        ]'
    },
    @{
        Title = "WF-2024-004: Travel Expense Policy"
        PolicyId = 104
        PolicyTitle = "Travel Expense Policy"
        TemplateId = 4
        Status = "Rejected"
        CurrentStage = 2
        TotalStages = 2
        InitiatedById = $users[5].Id
        InitiatedByName = $users[5].Name
        InitiatedDate = $now.AddDays(-14).ToString("o")
        DueDate = $now.AddDays(-4).ToString("o")
        CompletedDate = $now.AddDays(-6).ToString("o")
        FinalDecision = "Rejected"
        FinalComments = "Compliance concerns with international travel reimbursement section. Please revise section 4.2 and resubmit."
        IsUrgent = $false
        Priority = "Normal"
        StagesJson = '[
            {"stageNumber": 1, "stageName": "Finance Director Review", "status": "Approved", "completedDate": "' + $now.AddDays(-10).ToString("o") + '"},
            {"stageNumber": 2, "stageName": "Compliance Verification", "status": "Rejected", "completedDate": "' + $now.AddDays(-6).ToString("o") + '", "comments": "Compliance concerns with section 4.2"}
        ]'
    },
    @{
        Title = "WF-2024-005: Cybersecurity Incident Response"
        PolicyId = 105
        PolicyTitle = "Cybersecurity Incident Response Policy"
        TemplateId = 6
        Status = "InProgress"
        CurrentStage = 3
        TotalStages = 4
        InitiatedById = $users[4].Id
        InitiatedByName = $users[4].Name
        InitiatedDate = $now.AddDays(-10).ToString("o")
        DueDate = $now.AddDays(8).ToString("o")
        IsUrgent = $true
        Priority = "Critical"
        StagesJson = '[
            {"stageNumber": 1, "stageName": "Legal Draft Review", "status": "Approved", "completedDate": "' + $now.AddDays(-6).ToString("o") + '"},
            {"stageNumber": 2, "stageName": "Compliance Assessment", "status": "Approved", "completedDate": "' + $now.AddDays(-3).ToString("o") + '"},
            {"stageNumber": 3, "stageName": "HR Impact Review", "status": "Pending", "dueDate": "' + $now.AddDays(0).ToString("o") + '"},
            {"stageNumber": 4, "stageName": "Executive Sign-off", "status": "NotStarted"}
        ]'
    },
    @{
        Title = "WF-2024-006: Anti-Harassment Policy Update"
        PolicyId = 106
        PolicyTitle = "Anti-Harassment Policy Update"
        TemplateId = 2
        Status = "Escalated"
        CurrentStage = 2
        TotalStages = 3
        InitiatedById = $users[0].Id
        InitiatedByName = $users[0].Name
        InitiatedDate = $now.AddDays(-12).ToString("o")
        DueDate = $now.AddDays(-2).ToString("o")
        EscalatedDate = $now.AddDays(-1).ToString("o")
        EscalatedToId = $users[7].Id
        EscalatedToName = $users[7].Name
        IsUrgent = $true
        Priority = "High"
        StagesJson = '[
            {"stageNumber": 1, "stageName": "HR Review", "status": "Approved", "completedDate": "' + $now.AddDays(-9).ToString("o") + '"},
            {"stageNumber": 2, "stageName": "Legal Review", "status": "Escalated", "escalatedDate": "' + $now.AddDays(-1).ToString("o") + '"},
            {"stageNumber": 3, "stageName": "Executive Approval", "status": "NotStarted"}
        ]'
    },
    @{
        Title = "WF-2024-007: Social Media Policy"
        PolicyId = 107
        PolicyTitle = "Social Media Usage Policy"
        TemplateId = 5
        Status = "Pending"
        CurrentStage = 1
        TotalStages = 1
        InitiatedById = $users[1].Id
        InitiatedByName = $users[1].Name
        InitiatedDate = $now.AddDays(-1).ToString("o")
        DueDate = $now.AddDays(2).ToString("o")
        IsUrgent = $false
        Priority = "Low"
        StagesJson = '[
            {"stageNumber": 1, "stageName": "Policy Committee", "status": "Pending", "dueDate": "' + $now.AddDays(2).ToString("o") + '"}
        ]'
    }
)

$workflowCount = 0
foreach ($workflow in $workflows) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $workflow | Out-Null
            $workflowCount++
            Write-Host "  Added: $($workflow.Title)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($workflow.Title): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($workflow.Title)" -ForegroundColor Gray
    }
}
Write-Host "  Created $workflowCount workflows`n" -ForegroundColor Green

# ============================================================================
# SEED: JML_Policy_ApprovalDecisions
# ============================================================================

$listName = "JML_Policy_ApprovalDecisions"
Write-Host "Seeding: $listName..." -ForegroundColor White

$decisions = @(
    # WF-001 Decisions
    @{
        Title = "WF-001-S1: Department Review"
        WorkflowId = 1
        PolicyId = 101
        StageNumber = 1
        StageName = "Department Review"
        ApproverId = $users[1].Id
        ApproverName = $users[1].Name
        ApproverEmail = $users[1].Email
        Status = "Approved"
        Decision = "Approved"
        Comments = "Policy content is comprehensive and aligns with our department objectives. Approved."
        RequestedDate = $now.AddDays(-5).ToString("o")
        DueDate = $now.AddDays(0).ToString("o")
        DecisionDate = $now.AddDays(-3).ToString("o")
        IsRequired = $true
        CanDelegate = $true
        NotificationSent = $true
        TeamsCardSent = $true
    },
    @{
        Title = "WF-001-S2: Policy Committee - Sarah"
        WorkflowId = 1
        PolicyId = 101
        StageNumber = 2
        StageName = "Policy Committee Approval"
        ApproverId = $users[0].Id
        ApproverName = $users[0].Name
        ApproverEmail = $users[0].Email
        Status = "Pending"
        RequestedDate = $now.AddDays(-3).ToString("o")
        DueDate = $now.AddDays(4).ToString("o")
        IsRequired = $true
        CanDelegate = $true
        NotificationSent = $true
        TeamsCardSent = $true
        ReminderSentDate = $now.AddDays(-1).ToString("o")
    },
    @{
        Title = "WF-001-S2: Policy Committee - Emily"
        WorkflowId = 1
        PolicyId = 101
        StageNumber = 2
        StageName = "Policy Committee Approval"
        ApproverId = $users[2].Id
        ApproverName = $users[2].Name
        ApproverEmail = $users[2].Email
        Status = "Pending"
        RequestedDate = $now.AddDays(-3).ToString("o")
        DueDate = $now.AddDays(4).ToString("o")
        IsRequired = $true
        CanDelegate = $true
        NotificationSent = $true
        TeamsCardSent = $true
    },

    # WF-002 Decisions
    @{
        Title = "WF-002-S1: Technical Review - Amanda"
        WorkflowId = 2
        PolicyId = 102
        StageNumber = 1
        StageName = "Technical Review"
        ApproverId = $users[4].Id
        ApproverName = $users[4].Name
        ApproverEmail = $users[4].Email
        Status = "Pending"
        RequestedDate = $now.AddDays(-2).ToString("o")
        DueDate = $now.AddDays(8).ToString("o")
        IsRequired = $true
        CanDelegate = $true
        NotificationSent = $true
        TeamsCardSent = $true
    },
    @{
        Title = "WF-002-S1: Technical Review - Emily"
        WorkflowId = 2
        PolicyId = 102
        StageNumber = 1
        StageName = "Technical Review"
        ApproverId = $users[2].Id
        ApproverName = $users[2].Name
        ApproverEmail = $users[2].Email
        Status = "Approved"
        Decision = "Approved"
        Comments = "Compliance requirements are met. Technical implementation details are sound."
        RequestedDate = $now.AddDays(-2).ToString("o")
        DueDate = $now.AddDays(8).ToString("o")
        DecisionDate = $now.AddDays(-1).ToString("o")
        IsRequired = $true
        CanDelegate = $true
        NotificationSent = $true
        TeamsCardSent = $true
    },

    # WF-003 Decisions (Completed workflow)
    @{
        Title = "WF-003-S1: HR Review"
        WorkflowId = 3
        PolicyId = 103
        StageNumber = 1
        StageName = "HR Review"
        ApproverId = $users[0].Id
        ApproverName = $users[0].Name
        ApproverEmail = $users[0].Email
        Status = "Approved"
        Decision = "Approved"
        Comments = "Benefits package is competitive and well-structured."
        RequestedDate = $now.AddDays(-21).ToString("o")
        DueDate = $now.AddDays(-16).ToString("o")
        DecisionDate = $now.AddDays(-18).ToString("o")
        IsRequired = $true
        CanDelegate = $true
        NotificationSent = $true
        TeamsCardSent = $true
    },
    @{
        Title = "WF-003-S2: Legal Review"
        WorkflowId = 3
        PolicyId = 103
        StageNumber = 2
        StageName = "Legal Review"
        ApproverId = $users[3].Id
        ApproverName = $users[3].Name
        ApproverEmail = $users[3].Email
        Status = "Approved"
        Decision = "Approved"
        Comments = "No legal concerns. Policy is compliant with employment law."
        RequestedDate = $now.AddDays(-18).ToString("o")
        DueDate = $now.AddDays(-11).ToString("o")
        DecisionDate = $now.AddDays(-12).ToString("o")
        IsRequired = $true
        CanDelegate = $false
        NotificationSent = $true
        TeamsCardSent = $true
    },
    @{
        Title = "WF-003-S3: Executive Approval"
        WorkflowId = 3
        PolicyId = 103
        StageNumber = 3
        StageName = "Executive Approval"
        ApproverId = $users[7].Id
        ApproverName = $users[7].Name
        ApproverEmail = $users[7].Email
        Status = "Approved"
        Decision = "Approved"
        Comments = "Excellent work. Approved for immediate publication."
        RequestedDate = $now.AddDays(-12).ToString("o")
        DueDate = $now.AddDays(-9).ToString("o")
        DecisionDate = $now.AddDays(-8).ToString("o")
        IsRequired = $true
        CanDelegate = $true
        NotificationSent = $true
        TeamsCardSent = $true
    },

    # WF-004 Decisions (Rejected workflow)
    @{
        Title = "WF-004-S1: Finance Director Review"
        WorkflowId = 4
        PolicyId = 104
        StageNumber = 1
        StageName = "Finance Director Review"
        ApproverId = $users[6].Id
        ApproverName = $users[6].Name
        ApproverEmail = $users[6].Email
        Status = "Approved"
        Decision = "Approved"
        Comments = "Budget allocations are appropriate. Approved from finance perspective."
        RequestedDate = $now.AddDays(-14).ToString("o")
        DueDate = $now.AddDays(-9).ToString("o")
        DecisionDate = $now.AddDays(-10).ToString("o")
        IsRequired = $true
        CanDelegate = $false
        NotificationSent = $true
        TeamsCardSent = $true
    },
    @{
        Title = "WF-004-S2: Compliance Verification"
        WorkflowId = 4
        PolicyId = 104
        StageNumber = 2
        StageName = "Compliance Verification"
        ApproverId = $users[2].Id
        ApproverName = $users[2].Name
        ApproverEmail = $users[2].Email
        Status = "Rejected"
        Decision = "Rejected"
        Comments = "Section 4.2 regarding international travel reimbursement has potential tax compliance issues. The per diem rates for EMEA region need to align with local tax regulations. Please consult with tax team and revise."
        RequestedDate = $now.AddDays(-10).ToString("o")
        DueDate = $now.AddDays(-5).ToString("o")
        DecisionDate = $now.AddDays(-6).ToString("o")
        IsRequired = $true
        CanDelegate = $false
        NotificationSent = $true
        TeamsCardSent = $true
    },

    # WF-005 Decisions (In Progress - Critical)
    @{
        Title = "WF-005-S1: Legal Draft Review"
        WorkflowId = 5
        PolicyId = 105
        StageNumber = 1
        StageName = "Legal Draft Review"
        ApproverId = $users[3].Id
        ApproverName = $users[3].Name
        ApproverEmail = $users[3].Email
        Status = "Approved"
        Decision = "Approved"
        Comments = "Legal framework is solid. Notification requirements meet regulatory standards."
        RequestedDate = $now.AddDays(-10).ToString("o")
        DueDate = $now.AddDays(-3).ToString("o")
        DecisionDate = $now.AddDays(-6).ToString("o")
        IsRequired = $true
        CanDelegate = $false
        NotificationSent = $true
        TeamsCardSent = $true
    },
    @{
        Title = "WF-005-S2: Compliance Assessment"
        WorkflowId = 5
        PolicyId = 105
        StageNumber = 2
        StageName = "Compliance Assessment"
        ApproverId = $users[2].Id
        ApproverName = $users[2].Name
        ApproverEmail = $users[2].Email
        Status = "Approved"
        Decision = "Approved"
        Comments = "Meets NIST and ISO 27001 requirements. Approved."
        RequestedDate = $now.AddDays(-6).ToString("o")
        DueDate = $now.AddDays(-1).ToString("o")
        DecisionDate = $now.AddDays(-3).ToString("o")
        IsRequired = $true
        CanDelegate = $true
        NotificationSent = $true
        TeamsCardSent = $true
    },
    @{
        Title = "WF-005-S3: HR Impact Review"
        WorkflowId = 5
        PolicyId = 105
        StageNumber = 3
        StageName = "HR Impact Review"
        ApproverId = $users[0].Id
        ApproverName = $users[0].Name
        ApproverEmail = $users[0].Email
        Status = "Pending"
        RequestedDate = $now.AddDays(-3).ToString("o")
        DueDate = $now.AddDays(0).ToString("o")
        IsRequired = $true
        CanDelegate = $true
        NotificationSent = $true
        TeamsCardSent = $true
        ReminderSentDate = $now.AddDays(-1).ToString("o")
    },

    # WF-006 Decisions (Escalated)
    @{
        Title = "WF-006-S1: HR Review"
        WorkflowId = 6
        PolicyId = 106
        StageNumber = 1
        StageName = "HR Review"
        ApproverId = $users[0].Id
        ApproverName = $users[0].Name
        ApproverEmail = $users[0].Email
        Status = "Approved"
        Decision = "Approved"
        Comments = "Critical policy update. HR fully supports this revision."
        RequestedDate = $now.AddDays(-12).ToString("o")
        DueDate = $now.AddDays(-7).ToString("o")
        DecisionDate = $now.AddDays(-9).ToString("o")
        IsRequired = $true
        CanDelegate = $true
        NotificationSent = $true
        TeamsCardSent = $true
    },
    @{
        Title = "WF-006-S2: Legal Review"
        WorkflowId = 6
        PolicyId = 106
        StageNumber = 2
        StageName = "Legal Review"
        ApproverId = $users[3].Id
        ApproverName = $users[3].Name
        ApproverEmail = $users[3].Email
        Status = "Escalated"
        RequestedDate = $now.AddDays(-9).ToString("o")
        DueDate = $now.AddDays(-2).ToString("o")
        EscalatedDate = $now.AddDays(-1).ToString("o")
        IsRequired = $true
        CanDelegate = $false
        NotificationSent = $true
        TeamsCardSent = $true
        ReminderSentDate = $now.AddDays(-3).ToString("o")
    }
)

$decisionCount = 0
foreach ($decision in $decisions) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $decision | Out-Null
            $decisionCount++
            Write-Host "  Added: $($decision.Title)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($decision.Title): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($decision.Title)" -ForegroundColor Gray
    }
}
Write-Host "  Created $decisionCount decisions`n" -ForegroundColor Green

# ============================================================================
# SEED: JML_Policy_ApprovalDelegations
# ============================================================================

$listName = "JML_Policy_ApprovalDelegations"
Write-Host "Seeding: $listName..." -ForegroundColor White

$delegations = @(
    @{
        Title = "DEL-001: Sarah Mitchell - Holiday Coverage"
        DelegatorId = $users[0].Id
        DelegatorName = $users[0].Name
        DelegatorEmail = $users[0].Email
        DelegateId = $users[1].Id
        DelegateName = $users[1].Name
        DelegateEmail = $users[1].Email
        DelegationType = "Temporary"
        Scope = "All"
        StartDate = $now.AddDays(14).ToString("o")
        EndDate = $now.AddDays(28).ToString("o")
        Reason = "Annual leave - Christmas holiday period. James Chen will handle all policy approvals during this time."
        IsActive = $true
        NotifyDelegator = $true
        NotifyDelegate = $true
    },
    @{
        Title = "DEL-002: Michael Thompson - Sabbatical"
        DelegatorId = $users[3].Id
        DelegatorName = $users[3].Name
        DelegatorEmail = $users[3].Email
        DelegateId = $users[2].Id
        DelegateName = $users[2].Name
        DelegateEmail = $users[2].Email
        DelegationType = "OutOfOffice"
        Scope = "Category"
        ScopeCategory = "Legal"
        StartDate = $now.AddDays(-30).ToString("o")
        EndDate = $now.AddDays(60).ToString("o")
        Reason = "3-month sabbatical. Emily Rodriguez (Compliance Officer) will handle legal policy reviews during this period."
        IsActive = $true
        NotifyDelegator = $true
        NotifyDelegate = $true
    },
    @{
        Title = "DEL-003: Jennifer Walsh - Maternity Leave (Expired)"
        DelegatorId = $users[6].Id
        DelegatorName = $users[6].Name
        DelegatorEmail = $users[6].Email
        DelegateId = $users[5].Id
        DelegateName = $users[5].Name
        DelegateEmail = $users[5].Email
        DelegationType = "Temporary"
        Scope = "Category"
        ScopeCategory = "Finance"
        StartDate = $now.AddDays(-90).ToString("o")
        EndDate = $now.AddDays(-10).ToString("o")
        Reason = "Maternity leave coverage. David Kim handled finance policy approvals."
        IsActive = $false
        RevokedDate = $now.AddDays(-10).ToString("o")
        NotifyDelegator = $true
        NotifyDelegate = $true
    },
    @{
        Title = "DEL-004: Amanda Foster - Conference"
        DelegatorId = $users[4].Id
        DelegatorName = $users[4].Name
        DelegatorEmail = $users[4].Email
        DelegateId = $users[5].Id
        DelegateName = $users[5].Name
        DelegateEmail = $users[5].Email
        DelegationType = "Temporary"
        Scope = "Category"
        ScopeCategory = "IT"
        StartDate = $now.AddDays(7).ToString("o")
        EndDate = $now.AddDays(12).ToString("o")
        Reason = "Attending RSA Conference in San Francisco. David will cover IT security policy approvals."
        IsActive = $true
        NotifyDelegator = $true
        NotifyDelegate = $true
    },
    @{
        Title = "DEL-005: Robert Garcia - Permanent Deputy"
        DelegatorId = $users[7].Id
        DelegatorName = $users[7].Name
        DelegatorEmail = $users[7].Email
        DelegateId = $users[0].Id
        DelegateName = $users[0].Name
        DelegateEmail = $users[0].Email
        DelegationType = "Permanent"
        Scope = "All"
        StartDate = $now.AddDays(-180).ToString("o")
        Reason = "Sarah Mitchell authorized as permanent deputy for executive policy approvals when CEO is unavailable."
        IsActive = $true
        NotifyDelegator = $false
        NotifyDelegate = $true
    }
)

$delegationCount = 0
foreach ($delegation in $delegations) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $delegation | Out-Null
            $delegationCount++
            Write-Host "  Added: $($delegation.Title)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($delegation.Title): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($delegation.Title)" -ForegroundColor Gray
    }
}
Write-Host "  Created $delegationCount delegations`n" -ForegroundColor Green

# ============================================================================
# SEED: JML_Policy_EscalationRules
# ============================================================================

$listName = "JML_Policy_EscalationRules"
Write-Host "Seeding: $listName..." -ForegroundColor White

$escalationRules = @(
    @{
        Title = "ESC-001: Standard Reminder (48 hours)"
        RuleName = "Standard 48-Hour Reminder"
        Description = "Send reminder notification when approval has been pending for 48 hours. Applies to all non-urgent policies."
        TriggerType = "HoursOverdue"
        TriggerValue = 48
        ActionType = "Notify"
        NotifyOriginalApprover = $true
        NotifyInitiator = $false
        NotifyPolicyOwner = $false
        CustomEmailSubject = "Reminder: Policy approval pending your review"
        CustomEmailBody = "This is a friendly reminder that you have a policy approval request pending your review. Please take action within the next 24 hours to avoid escalation."
        AppliesTo = "All"
        IsActive = $true
        Priority = 1
        MaxEscalations = 2
    },
    @{
        Title = "ESC-002: Manager Notification (72 hours)"
        RuleName = "72-Hour Manager Escalation"
        Description = "Notify manager when approval has been pending for 72 hours without action."
        TriggerType = "HoursOverdue"
        TriggerValue = 72
        ActionType = "NotifyManager"
        NotifyOriginalApprover = $true
        NotifyInitiator = $true
        NotifyPolicyOwner = $false
        CustomEmailSubject = "Escalation: Policy approval requires attention"
        CustomEmailBody = "An approval request assigned to your direct report has been pending for over 72 hours. Please follow up to ensure timely review."
        AppliesTo = "All"
        IsActive = $true
        Priority = 2
        MaxEscalations = 1
    },
    @{
        Title = "ESC-003: Critical Policy - Urgent Escalation"
        RuleName = "Critical Policy 24-Hour Escalation"
        Description = "Escalate critical/urgent policies to executive sponsor after 24 hours."
        TriggerType = "HoursOverdue"
        TriggerValue = 24
        ActionType = "Reassign"
        ActionTargetId = $users[7].Id
        ActionTargetName = $users[7].Name
        ActionTargetEmail = $users[7].Email
        NotifyOriginalApprover = $true
        NotifyInitiator = $true
        NotifyPolicyOwner = $true
        CustomEmailSubject = "URGENT: Critical policy requires immediate attention"
        CustomEmailBody = "A critical policy approval has exceeded the 24-hour SLA. This has been escalated to executive leadership for immediate action."
        AppliesTo = "Priority"
        AppliesToValue = "Critical"
        IsActive = $true
        Priority = 0
        MaxEscalations = 1
    },
    @{
        Title = "ESC-004: High Priority - 48 Hour Escalation"
        RuleName = "High Priority 48-Hour Escalation"
        Description = "Escalate high priority policies after 48 hours of no response."
        TriggerType = "HoursOverdue"
        TriggerValue = 48
        ActionType = "NotifyManager"
        NotifyOriginalApprover = $true
        NotifyInitiator = $true
        NotifyPolicyOwner = $true
        CustomEmailSubject = "Escalation: High priority policy approval overdue"
        CustomEmailBody = "A high priority policy approval has been pending for 48 hours. Immediate attention is required."
        AppliesTo = "Priority"
        AppliesToValue = "High"
        IsActive = $true
        Priority = 1
        MaxEscalations = 2
    },
    @{
        Title = "ESC-005: Auto-Approve Low Risk (120 hours)"
        RuleName = "Auto-Approve Low Risk After 5 Days"
        Description = "Automatically approve low-priority, low-risk policy updates if no response after 5 business days."
        TriggerType = "HoursOverdue"
        TriggerValue = 120
        ActionType = "AutoApprove"
        NotifyOriginalApprover = $true
        NotifyInitiator = $true
        NotifyPolicyOwner = $true
        CustomEmailSubject = "Policy Auto-Approved Due to Timeout"
        CustomEmailBody = "The following policy has been automatically approved after 5 business days without response. If you have concerns, please contact the policy owner."
        AppliesTo = "Priority"
        AppliesToValue = "Low"
        IsActive = $true
        Priority = 10
        MaxEscalations = 1
    },
    @{
        Title = "ESC-006: Legal Policy - Compliance Escalation"
        RuleName = "Legal Policy Compliance Escalation"
        Description = "Escalate legal policies to compliance officer if stuck for more than 3 days."
        TriggerType = "DaysOverdue"
        TriggerValue = 3
        ActionType = "Reassign"
        ActionTargetId = $users[2].Id
        ActionTargetName = $users[2].Name
        ActionTargetEmail = $users[2].Email
        NotifyOriginalApprover = $true
        NotifyInitiator = $true
        NotifyPolicyOwner = $true
        CustomEmailSubject = "Legal Policy Escalation to Compliance"
        CustomEmailBody = "A legal policy review has exceeded the 3-day SLA. This has been escalated to the Compliance Officer for review."
        AppliesTo = "Category"
        AppliesToValue = "Legal"
        IsActive = $true
        Priority = 2
        MaxEscalations = 1
    },
    @{
        Title = "ESC-007: HR Policy - Executive Notification"
        RuleName = "HR Policy Executive Alert"
        Description = "Notify executive sponsor when HR policies are delayed beyond 5 days."
        TriggerType = "DaysOverdue"
        TriggerValue = 5
        ActionType = "NotifyManager"
        NotifyOriginalApprover = $true
        NotifyInitiator = $true
        NotifyPolicyOwner = $true
        CustomEmailSubject = "HR Policy Approval Delay Alert"
        CustomEmailBody = "An HR policy approval has been delayed beyond the expected timeframe. Executive visibility has been enabled."
        AppliesTo = "Category"
        AppliesToValue = "HR"
        IsActive = $true
        Priority = 3
        MaxEscalations = 1
    },
    @{
        Title = "ESC-008: Finance Policy - Dual Control Timeout"
        RuleName = "Finance Dual Control Timeout"
        Description = "Auto-reject finance policies if dual control approval not completed within 7 days."
        TriggerType = "DaysOverdue"
        TriggerValue = 7
        ActionType = "AutoReject"
        NotifyOriginalApprover = $true
        NotifyInitiator = $true
        NotifyPolicyOwner = $true
        CustomEmailSubject = "Finance Policy Auto-Rejected - Resubmission Required"
        CustomEmailBody = "The finance policy approval has been automatically rejected due to dual control timeout. Please resubmit for review."
        AppliesTo = "Category"
        AppliesToValue = "Finance"
        IsActive = $true
        Priority = 5
        MaxEscalations = 1
    }
)

$escalationCount = 0
foreach ($rule in $escalationRules) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $rule | Out-Null
            $escalationCount++
            Write-Host "  Added: $($rule.RuleName)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($rule.RuleName): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($rule.RuleName)" -ForegroundColor Gray
    }
}
Write-Host "  Created $escalationCount escalation rules`n" -ForegroundColor Green

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Sample Data Seeding Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

Write-Host "`nData created:" -ForegroundColor White
Write-Host "  - $templateCount Approval Templates (various workflow patterns)" -ForegroundColor Gray
Write-Host "  - $workflowCount Workflow Instances (different statuses)" -ForegroundColor Gray
Write-Host "  - $decisionCount Approval Decisions (approvals, rejections, pending)" -ForegroundColor Gray
Write-Host "  - $delegationCount Delegation Records (active and expired)" -ForegroundColor Gray
Write-Host "  - $escalationCount Escalation Rules (various triggers and actions)" -ForegroundColor Gray

Write-Host "`nSample scenarios included:" -ForegroundColor Yellow
Write-Host "  - Completed approval workflow (Employee Benefits)" -ForegroundColor Gray
Write-Host "  - Rejected workflow with detailed feedback (Travel Expense)" -ForegroundColor Gray
Write-Host "  - Escalated workflow requiring attention (Anti-Harassment)" -ForegroundColor Gray
Write-Host "  - Critical in-progress workflow (Cybersecurity Incident Response)" -ForegroundColor Gray
Write-Host "  - Multi-stage approval patterns (2, 3, and 4 stage workflows)" -ForegroundColor Gray
Write-Host "  - Various delegation scenarios (holiday, sabbatical, permanent deputy)" -ForegroundColor Gray
Write-Host "  - Comprehensive escalation rules (reminders, auto-actions, reassignments)" -ForegroundColor Gray
