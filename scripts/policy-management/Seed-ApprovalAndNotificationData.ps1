# ============================================================================
# Seed-ApprovalAndNotificationData.ps1
# Seeds PM_Approvals, PM_ApprovalChains, PM_ApprovalHistory,
# PM_ApprovalDelegations, PM_ApprovalTemplates, PM_Notifications,
# and PM_NotificationQueue with realistic sample data.
#
# PREREQUISITE: You must already be connected to SharePoint via PnP PowerShell
# PREREQUISITE: Run 08-Approval-Lists.ps1 and 07-Notification-Lists.ps1 first
# ============================================================================

param(
    [switch]$WhatIf = $false
)

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Approval & Notification - Sample Data Seeding" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

if ($WhatIf) {
    Write-Host "[WhatIf Mode] No changes will be made`n" -ForegroundColor Yellow
}

# ============================================================================
# SAMPLE USERS
# ============================================================================

Write-Host "Fetching site users..." -ForegroundColor White
$siteUsers = Get-PnPUser | Where-Object { $_.Email -ne "" } | Select-Object -First 10

if ($siteUsers.Count -lt 3) {
    Write-Host "Warning: Need at least 3 users with emails. Found: $($siteUsers.Count)" -ForegroundColor Yellow
    Write-Host "Using placeholder user data" -ForegroundColor Yellow

    $users = @(
        @{ Id = 1; Name = "Sarah Mitchell"; Email = "sarah.mitchell@contoso.com" },
        @{ Id = 2; Name = "James Chen"; Email = "james.chen@contoso.com" },
        @{ Id = 3; Name = "Emily Rodriguez"; Email = "emily.rodriguez@contoso.com" },
        @{ Id = 4; Name = "Michael Thompson"; Email = "michael.thompson@contoso.com" },
        @{ Id = 5; Name = "Amanda Foster"; Email = "amanda.foster@contoso.com" },
        @{ Id = 6; Name = "David Kim"; Email = "david.kim@contoso.com" },
        @{ Id = 7; Name = "Jennifer Walsh"; Email = "jennifer.walsh@contoso.com" },
        @{ Id = 8; Name = "Robert Garcia"; Email = "robert.garcia@contoso.com" }
    )
} else {
    $users = $siteUsers | ForEach-Object {
        @{ Id = $_.Id; Name = $_.Title; Email = $_.Email }
    }
}

Write-Host "Using $($users.Count) users for sample data`n" -ForegroundColor Gray

$now = Get-Date

# ============================================================================
# SEED: PM_ApprovalTemplates
# ============================================================================

$listName = "PM_ApprovalTemplates"
Write-Host "Seeding: $listName..." -ForegroundColor White

$templates = @(
    @{
        Title = "Standard Policy Approval"
        Description = "Standard two-level sequential approval for general policies"
        ProcessTypes = '["Policy"]'
        ApprovalType = "Sequential"
        Levels = '[{"Level":1,"ApproverIds":[' + $users[1].Id + '],"ApprovalType":"Sequential","DueDays":5,"ReasonRequired":false,"AllowDelegation":true,"EscalateToManagerOnDelay":true},{"Level":2,"ApproverIds":[' + $users[0].Id + ',' + $users[2].Id + '],"ApprovalType":"Parallel","DueDays":7,"ReasonRequired":true,"AllowDelegation":true,"EscalateToManagerOnDelay":true}]'
        RequireComments = $true
        AllowDelegation = $true
        AutoEscalationDays = 3
        EscalationAction = "Notify"
        IsActive = $true
    },
    @{
        Title = "HR Policy - Three Level"
        Description = "Three-level approval for HR policies: HR Review, Legal Review, Executive Sign-off"
        ProcessTypes = '["Policy"]'
        ApprovalType = "Sequential"
        Levels = '[{"Level":1,"ApproverIds":[' + $users[0].Id + '],"ApprovalType":"Sequential","DueDays":5,"ReasonRequired":false,"AllowDelegation":true,"EscalateToManagerOnDelay":true},{"Level":2,"ApproverIds":[' + $users[3].Id + '],"ApprovalType":"Sequential","DueDays":7,"ReasonRequired":true,"AllowDelegation":false,"EscalateToManagerOnDelay":true},{"Level":3,"ApproverIds":[' + $users[7].Id + '],"ApprovalType":"Sequential","DueDays":3,"ReasonRequired":true,"AllowDelegation":true,"EscalateToManagerOnDelay":false}]'
        RequireComments = $true
        AllowDelegation = $true
        AutoEscalationDays = 5
        EscalationAction = "AssignToManager"
        IsActive = $true
    },
    @{
        Title = "Quick Approval - Single Level"
        Description = "Fast-track single approver for low-risk policy updates"
        ProcessTypes = '["Policy"]'
        ApprovalType = "FirstApprover"
        Levels = '[{"Level":1,"ApproverIds":[' + $users[0].Id + ',' + $users[1].Id + ',' + $users[2].Id + '],"ApprovalType":"FirstApprover","DueDays":3,"ReasonRequired":false,"AllowDelegation":true,"EscalateToManagerOnDelay":false}]'
        RequireComments = $false
        AllowDelegation = $true
        AutoEscalationDays = 2
        EscalationAction = "AutoApprove"
        IsActive = $true
    },
    @{
        Title = "IT Security Policy"
        Description = "Parallel approval by IT Security and Compliance, then management sign-off"
        ProcessTypes = '["Policy"]'
        ApprovalType = "Sequential"
        Levels = '[{"Level":1,"ApproverIds":[' + $users[4].Id + ',' + $users[2].Id + '],"ApprovalType":"Parallel","DueDays":10,"ReasonRequired":true,"AllowDelegation":true,"EscalateToManagerOnDelay":true},{"Level":2,"ApproverIds":[' + $users[5].Id + '],"ApprovalType":"Sequential","DueDays":5,"ReasonRequired":true,"AllowDelegation":true,"EscalateToManagerOnDelay":true}]'
        RequireComments = $true
        AllowDelegation = $true
        AutoEscalationDays = 7
        EscalationAction = "Notify"
        IsActive = $true
    }
)

$templateCount = 0
foreach ($t in $templates) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $t | Out-Null
            $templateCount++
            Write-Host "  Added: $($t.Title)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($t.Title): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($t.Title)" -ForegroundColor Gray
    }
}
Write-Host "  Created $templateCount templates`n" -ForegroundColor Green

# ============================================================================
# SEED: PM_ApprovalChains
# ============================================================================

$listName = "PM_ApprovalChains"
Write-Host "Seeding: $listName..." -ForegroundColor White

$chains = @(
    @{
        Title = "Remote Work Policy Approval"
        ProcessID = 101
        ChainName = "Standard Policy Approval"
        ApprovalType = "Sequential"
        Levels = $templates[0].Levels
        RequireComments = $true
        AllowDelegation = $true
        AutoEscalationDays = 3
        EscalationAction = "Notify"
        CurrentLevel = 2
        OverallStatus = "Pending"
        IsActive = $true
        StartDate = $now.AddDays(-5).ToString("o")
    },
    @{
        Title = "Data Protection Policy Approval"
        ProcessID = 102
        ChainName = "IT Security Policy"
        ApprovalType = "Sequential"
        Levels = $templates[3].Levels
        RequireComments = $true
        AllowDelegation = $true
        AutoEscalationDays = 7
        EscalationAction = "Notify"
        CurrentLevel = 1
        OverallStatus = "Pending"
        IsActive = $true
        StartDate = $now.AddDays(-2).ToString("o")
    },
    @{
        Title = "Employee Benefits Policy Approval"
        ProcessID = 103
        ChainName = "HR Policy - Three Level"
        ApprovalType = "Sequential"
        Levels = $templates[1].Levels
        RequireComments = $true
        AllowDelegation = $true
        AutoEscalationDays = 5
        EscalationAction = "AssignToManager"
        CurrentLevel = 3
        OverallStatus = "Approved"
        IsActive = $false
        StartDate = $now.AddDays(-21).ToString("o")
        CompletedDate = $now.AddDays(-8).ToString("o")
    },
    @{
        Title = "Travel Expense Policy Approval"
        ProcessID = 104
        ChainName = "Standard Policy Approval"
        ApprovalType = "Sequential"
        Levels = $templates[0].Levels
        RequireComments = $true
        AllowDelegation = $true
        AutoEscalationDays = 3
        EscalationAction = "Notify"
        CurrentLevel = 2
        OverallStatus = "Rejected"
        IsActive = $false
        StartDate = $now.AddDays(-14).ToString("o")
        CompletedDate = $now.AddDays(-6).ToString("o")
    },
    @{
        Title = "Social Media Policy Approval"
        ProcessID = 107
        ChainName = "Quick Approval - Single Level"
        ApprovalType = "FirstApprover"
        Levels = $templates[2].Levels
        RequireComments = $false
        AllowDelegation = $true
        AutoEscalationDays = 2
        EscalationAction = "AutoApprove"
        CurrentLevel = 1
        OverallStatus = "Pending"
        IsActive = $true
        StartDate = $now.AddDays(-1).ToString("o")
    }
)

$chainCount = 0
foreach ($c in $chains) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $c | Out-Null
            $chainCount++
            Write-Host "  Added: $($c.Title)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($c.Title): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($c.Title)" -ForegroundColor Gray
    }
}
Write-Host "  Created $chainCount chains`n" -ForegroundColor Green

# ============================================================================
# SEED: PM_Approvals
# ============================================================================

$listName = "PM_Approvals"
Write-Host "Seeding: $listName..." -ForegroundColor White

$approvals = @(
    # Chain 1: Remote Work Policy (Level 1 done, Level 2 pending)
    @{
        Title = "Remote Work - L1 - James Chen"
        ProcessID = 101
        ApprovalChainId = 1
        ApprovalLevel = 1
        ApprovalSequence = 1
        ApprovalType = "Sequential"
        Status = "Approved"
        ApproverId = $users[1].Id
        RequestedDate = $now.AddDays(-5).ToString("o")
        DueDate = $now.AddDays(0).ToString("o")
        CompletedDate = $now.AddDays(-3).ToString("o")
        ResponseTime = 48
        Decision = "Approved"
        Comments = "Policy content is comprehensive and aligns with our department objectives."
        ReasonRequired = $false
        IsOverdue = $false
        EscalationLevel = 0
    },
    @{
        Title = "Remote Work - L2 - Sarah Mitchell"
        ProcessID = 101
        ApprovalChainId = 1
        ApprovalLevel = 2
        ApprovalSequence = 1
        ApprovalType = "Parallel"
        Status = "Pending"
        ApproverId = $users[0].Id
        RequestedDate = $now.AddDays(-3).ToString("o")
        DueDate = $now.AddDays(4).ToString("o")
        ReasonRequired = $true
        IsOverdue = $false
        EscalationLevel = 0
    },
    @{
        Title = "Remote Work - L2 - Emily Rodriguez"
        ProcessID = 101
        ApprovalChainId = 1
        ApprovalLevel = 2
        ApprovalSequence = 2
        ApprovalType = "Parallel"
        Status = "Pending"
        ApproverId = $users[2].Id
        RequestedDate = $now.AddDays(-3).ToString("o")
        DueDate = $now.AddDays(4).ToString("o")
        ReasonRequired = $true
        IsOverdue = $false
        EscalationLevel = 0
    },

    # Chain 2: Data Protection Policy (Level 1 in progress)
    @{
        Title = "Data Protection - L1 - Amanda Foster"
        ProcessID = 102
        ApprovalChainId = 2
        ApprovalLevel = 1
        ApprovalSequence = 1
        ApprovalType = "Parallel"
        Status = "Pending"
        ApproverId = $users[4].Id
        RequestedDate = $now.AddDays(-2).ToString("o")
        DueDate = $now.AddDays(8).ToString("o")
        ReasonRequired = $true
        IsOverdue = $false
        EscalationLevel = 0
    },
    @{
        Title = "Data Protection - L1 - Emily Rodriguez"
        ProcessID = 102
        ApprovalChainId = 2
        ApprovalLevel = 1
        ApprovalSequence = 2
        ApprovalType = "Parallel"
        Status = "Approved"
        ApproverId = $users[2].Id
        RequestedDate = $now.AddDays(-2).ToString("o")
        DueDate = $now.AddDays(8).ToString("o")
        CompletedDate = $now.AddDays(-1).ToString("o")
        ResponseTime = 24
        Decision = "Approved"
        Comments = "Compliance requirements are met. Technical implementation details are sound."
        ReasonRequired = $true
        IsOverdue = $false
        EscalationLevel = 0
    },

    # Chain 3: Employee Benefits (All approved)
    @{
        Title = "Employee Benefits - L1 - Sarah Mitchell"
        ProcessID = 103
        ApprovalChainId = 3
        ApprovalLevel = 1
        ApprovalSequence = 1
        ApprovalType = "Sequential"
        Status = "Approved"
        ApproverId = $users[0].Id
        RequestedDate = $now.AddDays(-21).ToString("o")
        DueDate = $now.AddDays(-16).ToString("o")
        CompletedDate = $now.AddDays(-18).ToString("o")
        ResponseTime = 72
        Decision = "Approved"
        Comments = "Benefits package is competitive and well-structured."
        ReasonRequired = $false
        IsOverdue = $false
        EscalationLevel = 0
    },
    @{
        Title = "Employee Benefits - L2 - Michael Thompson"
        ProcessID = 103
        ApprovalChainId = 3
        ApprovalLevel = 2
        ApprovalSequence = 1
        ApprovalType = "Sequential"
        Status = "Approved"
        ApproverId = $users[3].Id
        RequestedDate = $now.AddDays(-18).ToString("o")
        DueDate = $now.AddDays(-11).ToString("o")
        CompletedDate = $now.AddDays(-12).ToString("o")
        ResponseTime = 144
        Decision = "Approved"
        Comments = "No legal concerns. Policy is compliant with employment law."
        ReasonRequired = $true
        IsOverdue = $false
        EscalationLevel = 0
    },
    @{
        Title = "Employee Benefits - L3 - Robert Garcia"
        ProcessID = 103
        ApprovalChainId = 3
        ApprovalLevel = 3
        ApprovalSequence = 1
        ApprovalType = "Sequential"
        Status = "Approved"
        ApproverId = $users[7].Id
        RequestedDate = $now.AddDays(-12).ToString("o")
        DueDate = $now.AddDays(-9).ToString("o")
        CompletedDate = $now.AddDays(-8).ToString("o")
        ResponseTime = 96
        Decision = "Approved"
        Comments = "Excellent work. Approved for immediate publication."
        ReasonRequired = $true
        IsOverdue = $false
        EscalationLevel = 0
    },

    # Chain 4: Travel Expense (Rejected)
    @{
        Title = "Travel Expense - L1 - James Chen"
        ProcessID = 104
        ApprovalChainId = 4
        ApprovalLevel = 1
        ApprovalSequence = 1
        ApprovalType = "Sequential"
        Status = "Approved"
        ApproverId = $users[1].Id
        RequestedDate = $now.AddDays(-14).ToString("o")
        DueDate = $now.AddDays(-9).ToString("o")
        CompletedDate = $now.AddDays(-10).ToString("o")
        ResponseTime = 96
        Decision = "Approved"
        Comments = "Budget allocations are appropriate."
        ReasonRequired = $false
        IsOverdue = $false
        EscalationLevel = 0
    },
    @{
        Title = "Travel Expense - L2 - Emily Rodriguez"
        ProcessID = 104
        ApprovalChainId = 4
        ApprovalLevel = 2
        ApprovalSequence = 1
        ApprovalType = "Parallel"
        Status = "Rejected"
        ApproverId = $users[2].Id
        RequestedDate = $now.AddDays(-10).ToString("o")
        DueDate = $now.AddDays(-3).ToString("o")
        CompletedDate = $now.AddDays(-6).ToString("o")
        ResponseTime = 96
        Decision = "Rejected"
        Comments = "Section 4.2 regarding international travel reimbursement has potential tax compliance issues. The per diem rates for EMEA region need to align with local tax regulations."
        ReasonRequired = $true
        IsOverdue = $false
        EscalationLevel = 0
    },

    # Chain 5: Social Media Policy (Pending quick approval)
    @{
        Title = "Social Media - L1 - Sarah Mitchell"
        ProcessID = 107
        ApprovalChainId = 5
        ApprovalLevel = 1
        ApprovalSequence = 1
        ApprovalType = "FirstApprover"
        Status = "Pending"
        ApproverId = $users[0].Id
        RequestedDate = $now.AddDays(-1).ToString("o")
        DueDate = $now.AddDays(2).ToString("o")
        ReasonRequired = $false
        IsOverdue = $false
        EscalationLevel = 0
    },
    @{
        Title = "Social Media - L1 - James Chen"
        ProcessID = 107
        ApprovalChainId = 5
        ApprovalLevel = 1
        ApprovalSequence = 2
        ApprovalType = "FirstApprover"
        Status = "Pending"
        ApproverId = $users[1].Id
        RequestedDate = $now.AddDays(-1).ToString("o")
        DueDate = $now.AddDays(2).ToString("o")
        ReasonRequired = $false
        IsOverdue = $false
        EscalationLevel = 0
    },
    @{
        Title = "Social Media - L1 - Emily Rodriguez"
        ProcessID = 107
        ApprovalChainId = 5
        ApprovalLevel = 1
        ApprovalSequence = 3
        ApprovalType = "FirstApprover"
        Status = "Pending"
        ApproverId = $users[2].Id
        RequestedDate = $now.AddDays(-1).ToString("o")
        DueDate = $now.AddDays(2).ToString("o")
        ReasonRequired = $false
        IsOverdue = $false
        EscalationLevel = 0
    }
)

$approvalCount = 0
foreach ($a in $approvals) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $a | Out-Null
            $approvalCount++
            Write-Host "  Added: $($a.Title)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($a.Title): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($a.Title)" -ForegroundColor Gray
    }
}
Write-Host "  Created $approvalCount approval records`n" -ForegroundColor Green

# ============================================================================
# SEED: PM_ApprovalHistory
# ============================================================================

$listName = "PM_ApprovalHistory"
Write-Host "Seeding: $listName..." -ForegroundColor White

$history = @(
    @{
        Title = "Remote Work - L1 Approved"
        ApprovalId = 1
        ProcessID = 101
        Action = "Approved"
        PerformedById = $users[1].Id
        ActionDate = $now.AddDays(-3).ToString("o")
        Comments = "Policy content is comprehensive and aligns with our department objectives."
        PreviousStatus = "Pending"
        NewStatus = "Approved"
    },
    @{
        Title = "Data Protection - L1 Approved (Emily)"
        ApprovalId = 5
        ProcessID = 102
        Action = "Approved"
        PerformedById = $users[2].Id
        ActionDate = $now.AddDays(-1).ToString("o")
        Comments = "Compliance requirements are met."
        PreviousStatus = "Pending"
        NewStatus = "Approved"
    },
    @{
        Title = "Employee Benefits - L1 Approved"
        ApprovalId = 6
        ProcessID = 103
        Action = "Approved"
        PerformedById = $users[0].Id
        ActionDate = $now.AddDays(-18).ToString("o")
        Comments = "Benefits package is competitive and well-structured."
        PreviousStatus = "Pending"
        NewStatus = "Approved"
    },
    @{
        Title = "Employee Benefits - L2 Approved"
        ApprovalId = 7
        ProcessID = 103
        Action = "Approved"
        PerformedById = $users[3].Id
        ActionDate = $now.AddDays(-12).ToString("o")
        Comments = "No legal concerns."
        PreviousStatus = "Pending"
        NewStatus = "Approved"
    },
    @{
        Title = "Employee Benefits - L3 Approved"
        ApprovalId = 8
        ProcessID = 103
        Action = "Approved"
        PerformedById = $users[7].Id
        ActionDate = $now.AddDays(-8).ToString("o")
        Comments = "Approved for immediate publication."
        PreviousStatus = "Pending"
        NewStatus = "Approved"
    },
    @{
        Title = "Travel Expense - L1 Approved"
        ApprovalId = 9
        ProcessID = 104
        Action = "Approved"
        PerformedById = $users[1].Id
        ActionDate = $now.AddDays(-10).ToString("o")
        Comments = "Budget allocations are appropriate."
        PreviousStatus = "Pending"
        NewStatus = "Approved"
    },
    @{
        Title = "Travel Expense - L2 Rejected"
        ApprovalId = 10
        ProcessID = 104
        Action = "Rejected"
        PerformedById = $users[2].Id
        ActionDate = $now.AddDays(-6).ToString("o")
        Comments = "Section 4.2 has tax compliance issues."
        PreviousStatus = "Pending"
        NewStatus = "Rejected"
    }
)

$historyCount = 0
foreach ($h in $history) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $h | Out-Null
            $historyCount++
            Write-Host "  Added: $($h.Title)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($h.Title): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($h.Title)" -ForegroundColor Gray
    }
}
Write-Host "  Created $historyCount history records`n" -ForegroundColor Green

# ============================================================================
# SEED: PM_ApprovalDelegations
# ============================================================================

$listName = "PM_ApprovalDelegations"
Write-Host "Seeding: $listName..." -ForegroundColor White

$delegations = @(
    @{
        Title = "Sarah Mitchell - Holiday Coverage"
        DelegatedById = $users[0].Id
        DelegatedToId = $users[1].Id
        StartDate = $now.AddDays(14).ToString("o")
        EndDate = $now.AddDays(28).ToString("o")
        IsActive = $true
        Reason = "Annual leave. James Chen will handle all policy approvals."
        AutoDelegate = $true
    },
    @{
        Title = "Michael Thompson - Sabbatical"
        DelegatedById = $users[3].Id
        DelegatedToId = $users[2].Id
        StartDate = $now.AddDays(-30).ToString("o")
        EndDate = $now.AddDays(60).ToString("o")
        IsActive = $true
        Reason = "3-month sabbatical. Emily Rodriguez handling legal policy reviews."
        AutoDelegate = $true
    },
    @{
        Title = "Robert Garcia - Permanent Deputy"
        DelegatedById = $users[7].Id
        DelegatedToId = $users[0].Id
        StartDate = $now.AddDays(-180).ToString("o")
        EndDate = $now.AddDays(365).ToString("o")
        IsActive = $true
        Reason = "Sarah Mitchell authorized as permanent deputy for executive approvals."
        AutoDelegate = $true
    }
)

$delegationCount = 0
foreach ($d in $delegations) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $d | Out-Null
            $delegationCount++
            Write-Host "  Added: $($d.Title)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($d.Title): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($d.Title)" -ForegroundColor Gray
    }
}
Write-Host "  Created $delegationCount delegations`n" -ForegroundColor Green

# ============================================================================
# SEED: PM_Notifications (In-App Notifications)
# ============================================================================

$listName = "PM_Notifications"
Write-Host "Seeding: $listName..." -ForegroundColor White

$notifications = @(
    @{
        Title = "New policy requires your approval"
        Message = "The Remote Work Policy v2.0 has been submitted for your approval. Please review and provide your decision by " + $now.AddDays(4).ToString("dd MMM yyyy") + "."
        RecipientId = $users[0].Id
        Type = "PolicyAcknowledgment"
        Priority = "High"
        IsRead = $false
        RelatedItemType = "Approval"
        RelatedItemId = 2
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=101"
    },
    @{
        Title = "New policy requires your approval"
        Message = "The Remote Work Policy v2.0 has been submitted for your approval. Please review and provide your decision by " + $now.AddDays(4).ToString("dd MMM yyyy") + "."
        RecipientId = $users[2].Id
        Type = "PolicyAcknowledgment"
        Priority = "High"
        IsRead = $false
        RelatedItemType = "Approval"
        RelatedItemId = 3
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=101"
    },
    @{
        Title = "Policy approved: Employee Benefits Policy"
        Message = "The Employee Benefits Policy has been fully approved and is now ready for publication. All three approval levels have been completed."
        RecipientId = $users[0].Id
        Type = "PolicyUpdate"
        Priority = "Normal"
        IsRead = $true
        RelatedItemType = "Policy"
        RelatedItemId = 103
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=103"
    },
    @{
        Title = "Policy rejected: Travel Expense Policy"
        Message = "The Travel Expense Policy has been rejected by Emily Rodriguez (Compliance Officer). Reason: Section 4.2 has tax compliance issues with EMEA per diem rates."
        RecipientId = $users[5].Id
        Type = "PolicyUpdate"
        Priority = "High"
        IsRead = $false
        RelatedItemType = "Policy"
        RelatedItemId = 104
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=104"
    },
    @{
        Title = "Approval reminder: Data Protection Policy"
        Message = "You have a pending approval for the Data Protection Policy. This is marked as urgent. Please review at your earliest convenience."
        RecipientId = $users[4].Id
        Type = "PolicyAcknowledgment"
        Priority = "Urgent"
        IsRead = $false
        RelatedItemType = "Approval"
        RelatedItemId = 4
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=102"
    },
    @{
        Title = "New policy published: Code of Conduct v3.1"
        Message = "A new version of the Code of Conduct policy has been published. Please read and acknowledge by " + $now.AddDays(14).ToString("dd MMM yyyy") + "."
        RecipientId = $users[1].Id
        Type = "Policy"
        Priority = "Normal"
        IsRead = $true
        RelatedItemType = "Policy"
        RelatedItemId = 110
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=110"
    },
    @{
        Title = "Delegation activated"
        Message = "Your delegation to James Chen has been activated. James will handle your approvals from " + $now.AddDays(14).ToString("dd MMM yyyy") + " to " + $now.AddDays(28).ToString("dd MMM yyyy") + "."
        RecipientId = $users[0].Id
        Type = "PolicyUpdate"
        Priority = "Normal"
        IsRead = $false
        RelatedItemType = "Delegation"
        RelatedItemId = 1
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyAdmin.aspx"
    },
    @{
        Title = "Policy expiring soon: Anti-Harassment Policy"
        Message = "The Anti-Harassment Policy is due for review within 30 days. As the policy owner, please initiate the review process."
        RecipientId = $users[0].Id
        Type = "PolicyExpiring"
        Priority = "High"
        IsRead = $false
        RelatedItemType = "Policy"
        RelatedItemId = 106
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=106"
    },
    @{
        Title = "Policy shared with you"
        Message = "James Chen shared the Cybersecurity Incident Response Policy with you. Click to view."
        RecipientId = $users[5].Id
        Type = "PolicyShare"
        Priority = "Low"
        IsRead = $true
        RelatedItemType = "Policy"
        RelatedItemId = 105
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=105"
    },
    @{
        Title = "New policy requires your acknowledgement"
        Message = "The updated IT Acceptable Use Policy requires your acknowledgement. Due date: " + $now.AddDays(7).ToString("dd MMM yyyy") + "."
        RecipientId = $users[3].Id
        Type = "PolicyAcknowledgment"
        Priority = "Normal"
        IsRead = $false
        RelatedItemType = "Policy"
        RelatedItemId = 108
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=108"
    },
    @{
        Title = "Social Media Policy pending approval"
        Message = "The Social Media Usage Policy has been submitted for quick approval. Any committee member can approve."
        RecipientId = $users[0].Id
        Type = "PolicyAcknowledgment"
        Priority = "Low"
        IsRead = $false
        RelatedItemType = "Approval"
        RelatedItemId = 11
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyDetails.aspx?policyId=107"
    },
    @{
        Title = "Weekly policy compliance summary"
        Message = "Your department has 92% policy compliance this week. 3 policies are pending acknowledgement from team members."
        RecipientId = $users[5].Id
        Type = "PolicyUpdate"
        Priority = "Low"
        IsRead = $true
        RelatedItemType = "Policy"
        RelatedItemId = 0
        ActionUrl = "/sites/PolicyManager/SitePages/PolicyHub.aspx"
    }
)

$notificationCount = 0
foreach ($n in $notifications) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $n | Out-Null
            $notificationCount++
            Write-Host "  Added: $($n.Title)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($n.Title): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($n.Title)" -ForegroundColor Gray
    }
}
Write-Host "  Created $notificationCount notifications`n" -ForegroundColor Green

# ============================================================================
# SEED: PM_NotificationQueue (Outbound notification queue)
# ============================================================================

$listName = "PM_NotificationQueue"
Write-Host "Seeding: $listName..." -ForegroundColor White

$queue = @(
    @{
        Title = "Approval Request - Remote Work Policy"
        NotificationType = "PolicyAcknowledgmentRequired"
        RecipientEmail = $users[0].Email
        RecipientUserId = $users[0].Id
        RecipientName = $users[0].Name
        SenderEmail = $users[1].Email
        SenderUserId = $users[1].Id
        SenderName = $users[1].Name
        PolicyId = 101
        PolicyTitle = "Remote Work Policy v2.0"
        PolicyVersion = "2.0"
        Message = "Your approval is required for the Remote Work Policy v2.0. Please review and approve by " + $now.AddDays(4).ToString("dd MMM yyyy") + "."
        Channel = "Email"
        Priority = "High"
        Status = "Sent"
        RetryCount = 0
        MaxRetries = 3
        SentTime = $now.AddDays(-3).ToString("o")
    },
    @{
        Title = "Approval Request - Remote Work Policy (Teams)"
        NotificationType = "PolicyAcknowledgmentRequired"
        RecipientEmail = $users[0].Email
        RecipientUserId = $users[0].Id
        RecipientName = $users[0].Name
        SenderEmail = $users[1].Email
        SenderUserId = $users[1].Id
        SenderName = $users[1].Name
        PolicyId = 101
        PolicyTitle = "Remote Work Policy v2.0"
        PolicyVersion = "2.0"
        Message = "Your approval is required for the Remote Work Policy v2.0."
        Channel = "Teams"
        Priority = "High"
        Status = "Sent"
        RetryCount = 0
        MaxRetries = 3
        SentTime = $now.AddDays(-3).ToString("o")
    },
    @{
        Title = "Policy Rejected Notification"
        NotificationType = "PolicyUpdated"
        RecipientEmail = $users[5].Email
        RecipientUserId = $users[5].Id
        RecipientName = $users[5].Name
        SenderEmail = "system@contoso.com"
        SenderName = "Policy Manager"
        PolicyId = 104
        PolicyTitle = "Travel Expense Policy"
        PolicyVersion = "1.0"
        Message = "The Travel Expense Policy has been rejected. Reason: Section 4.2 has tax compliance issues."
        Channel = "Email"
        Priority = "High"
        Status = "Sent"
        RetryCount = 0
        MaxRetries = 3
        SentTime = $now.AddDays(-6).ToString("o")
    },
    @{
        Title = "Reminder - Data Protection Approval"
        NotificationType = "PolicyAcknowledgmentRequired"
        RecipientEmail = $users[4].Email
        RecipientUserId = $users[4].Id
        RecipientName = $users[4].Name
        SenderEmail = "system@contoso.com"
        SenderName = "Policy Manager"
        PolicyId = 102
        PolicyTitle = "Data Protection Policy"
        PolicyVersion = "1.0"
        Message = "Reminder: Your approval for the Data Protection Policy is still pending. This is marked as urgent."
        Channel = "All"
        Priority = "Urgent"
        Status = "Pending"
        RetryCount = 0
        MaxRetries = 3
        ScheduledSendTime = $now.AddHours(1).ToString("o")
    },
    @{
        Title = "Policy Published - Code of Conduct"
        NotificationType = "PolicyPublished"
        RecipientEmail = $users[1].Email
        RecipientUserId = $users[1].Id
        RecipientName = $users[1].Name
        SenderEmail = "system@contoso.com"
        SenderName = "Policy Manager"
        PolicyId = 110
        PolicyTitle = "Code of Conduct v3.1"
        PolicyVersion = "3.1"
        Message = "A new version of the Code of Conduct has been published. Please acknowledge."
        Channel = "Email"
        Priority = "Normal"
        Status = "Sent"
        RetryCount = 0
        MaxRetries = 3
        SentTime = $now.AddDays(-2).ToString("o")
    },
    @{
        Title = "Failed notification - retry"
        NotificationType = "PolicyExpiring"
        RecipientEmail = "invalid@contoso.com"
        RecipientUserId = 999
        RecipientName = "Unknown User"
        SenderEmail = "system@contoso.com"
        SenderName = "Policy Manager"
        PolicyId = 106
        PolicyTitle = "Anti-Harassment Policy"
        PolicyVersion = "2.5"
        Message = "The Anti-Harassment Policy is due for review within 30 days."
        Channel = "Email"
        Priority = "Normal"
        Status = "Failed"
        RetryCount = 3
        MaxRetries = 3
        LastError = "Recipient email address not found in directory. Message delivery failed after 3 retries."
    }
)

$queueCount = 0
foreach ($q in $queue) {
    if (-not $WhatIf) {
        try {
            Add-PnPListItem -List $listName -Values $q | Out-Null
            $queueCount++
            Write-Host "  Added: $($q.Title)" -ForegroundColor Gray
        } catch {
            Write-Host "  Error adding $($q.Title): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WhatIf] Would add: $($q.Title)" -ForegroundColor Gray
    }
}
Write-Host "  Created $queueCount queue records`n" -ForegroundColor Green

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Sample Data Seeding Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

Write-Host "`nData created:" -ForegroundColor White
Write-Host "  - $templateCount Approval Templates" -ForegroundColor Gray
Write-Host "  - $chainCount Approval Chains (active workflows)" -ForegroundColor Gray
Write-Host "  - $approvalCount Approval Records (individual decisions)" -ForegroundColor Gray
Write-Host "  - $historyCount History Records (audit trail)" -ForegroundColor Gray
Write-Host "  - $delegationCount Delegation Records" -ForegroundColor Gray
Write-Host "  - $notificationCount In-App Notifications" -ForegroundColor Gray
Write-Host "  - $queueCount Notification Queue Items" -ForegroundColor Gray

Write-Host "`nSample scenarios:" -ForegroundColor Yellow
Write-Host "  - Remote Work Policy: Level 1 approved, Level 2 pending (2 parallel approvers)" -ForegroundColor Gray
Write-Host "  - Data Protection Policy: Urgent, Level 1 partially approved (1 of 2)" -ForegroundColor Gray
Write-Host "  - Employee Benefits: Fully approved through all 3 levels" -ForegroundColor Gray
Write-Host "  - Travel Expense: Rejected at Level 2 with compliance feedback" -ForegroundColor Gray
Write-Host "  - Social Media: Quick approval pending (any of 3 can approve)" -ForegroundColor Gray
Write-Host "  - Mix of read/unread notifications, sent/pending/failed queue items" -ForegroundColor Gray
