# ============================================================================
# Provision-PolicyApprovalLists.ps1
# Creates SharePoint lists for Policy Approval Workflow (Phase 2)
#
# PREREQUISITE: You must already be connected to SharePoint via PnP PowerShell
# ============================================================================

param(
    [switch]$WhatIf = $false
)

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Policy Approval Workflow - List Provisioning" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

if ($WhatIf) {
    Write-Host "[WhatIf Mode] No changes will be made`n" -ForegroundColor Yellow
}

# ============================================================================
# LIST 1: JML_Policy_ApprovalTemplates
# Stores reusable workflow templates
# ============================================================================

$listName = "JML_Policy_ApprovalTemplates"
Write-Host "Creating list: $listName..." -ForegroundColor White

$existingList = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($existingList) {
    Write-Host "  List already exists, skipping creation" -ForegroundColor Yellow
} else {
    if (-not $WhatIf) {
        New-PnPList -Title $listName -Template GenericList -EnableVersioning
        Write-Host "  List created" -ForegroundColor Green
    } else {
        Write-Host "  [WhatIf] Would create list" -ForegroundColor Gray
    }
}

# Add columns to ApprovalTemplates
$templateColumns = @(
    @{ Name = "TemplateName"; Type = "Text"; Required = $true },
    @{ Name = "Description"; Type = "Note"; Required = $false },
    @{ Name = "Category"; Type = "Choice"; Choices = @("HR", "IT", "Finance", "Legal", "Operations", "Compliance", "General") },
    @{ Name = "IsDefault"; Type = "Boolean" },
    @{ Name = "IsActive"; Type = "Boolean" },
    @{ Name = "StagesJson"; Type = "Note"; Required = $true },  # JSON array of stages
    @{ Name = "RequireAllStages"; Type = "Boolean" },
    @{ Name = "AllowParallelApproval"; Type = "Boolean" },
    @{ Name = "NotifyOnComplete"; Type = "Boolean" },
    @{ Name = "AutoArchiveOnComplete"; Type = "Boolean" },
    @{ Name = "CreatedById"; Type = "Number" },
    @{ Name = "ModifiedById"; Type = "Number" }
)

foreach ($col in $templateColumns) {
    $existingField = Get-PnPField -List $listName -Identity $col.Name -ErrorAction SilentlyContinue
    if (-not $existingField) {
        if (-not $WhatIf) {
            switch ($col.Type) {
                "Text" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Text -Required:$col.Required -AddToDefaultView
                }
                "Note" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Note -Required:$col.Required
                }
                "Choice" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Choice -Choices $col.Choices -AddToDefaultView
                }
                "Boolean" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Boolean -AddToDefaultView
                }
                "Number" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Number
                }
            }
            Write-Host "    Added column: $($col.Name)" -ForegroundColor Gray
        }
    } else {
        Write-Host "    Column exists: $($col.Name)" -ForegroundColor DarkGray
    }
}

Write-Host "  Completed: $listName`n" -ForegroundColor Green

# ============================================================================
# LIST 2: JML_Policy_ApprovalWorkflows
# Stores active workflow instances
# ============================================================================

$listName = "JML_Policy_ApprovalWorkflows"
Write-Host "Creating list: $listName..." -ForegroundColor White

$existingList = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($existingList) {
    Write-Host "  List already exists, skipping creation" -ForegroundColor Yellow
} else {
    if (-not $WhatIf) {
        New-PnPList -Title $listName -Template GenericList -EnableVersioning
        Write-Host "  List created" -ForegroundColor Green
    } else {
        Write-Host "  [WhatIf] Would create list" -ForegroundColor Gray
    }
}

# Add columns to ApprovalWorkflows
$workflowColumns = @(
    @{ Name = "PolicyId"; Type = "Number"; Required = $true; Indexed = $true },
    @{ Name = "PolicyTitle"; Type = "Text" },
    @{ Name = "TemplateId"; Type = "Number"; Indexed = $true },
    @{ Name = "Status"; Type = "Choice"; Choices = @("Pending", "InProgress", "Approved", "Rejected", "Cancelled", "Escalated") },
    @{ Name = "CurrentStage"; Type = "Number" },
    @{ Name = "TotalStages"; Type = "Number" },
    @{ Name = "StagesJson"; Type = "Note" },  # JSON array of stage status
    @{ Name = "InitiatedById"; Type = "Number"; Indexed = $true },
    @{ Name = "InitiatedByName"; Type = "Text" },
    @{ Name = "InitiatedDate"; Type = "DateTime" },
    @{ Name = "CompletedDate"; Type = "DateTime" },
    @{ Name = "DueDate"; Type = "DateTime"; Indexed = $true },
    @{ Name = "FinalDecision"; Type = "Choice"; Choices = @("Approved", "Rejected", "Cancelled") },
    @{ Name = "FinalComments"; Type = "Note" },
    @{ Name = "EscalatedDate"; Type = "DateTime" },
    @{ Name = "EscalatedToId"; Type = "Number" },
    @{ Name = "EscalatedToName"; Type = "Text" },
    @{ Name = "IsUrgent"; Type = "Boolean" },
    @{ Name = "Priority"; Type = "Choice"; Choices = @("Low", "Normal", "High", "Critical") }
)

foreach ($col in $workflowColumns) {
    $existingField = Get-PnPField -List $listName -Identity $col.Name -ErrorAction SilentlyContinue
    if (-not $existingField) {
        if (-not $WhatIf) {
            switch ($col.Type) {
                "Text" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Text -AddToDefaultView
                }
                "Note" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Note
                }
                "Choice" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Choice -Choices $col.Choices -AddToDefaultView
                }
                "Boolean" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Boolean
                }
                "Number" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Number -AddToDefaultView
                }
                "DateTime" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type DateTime -AddToDefaultView
                }
            }

            # Add index if specified
            if ($col.Indexed -and $col.Type -ne "Note") {
                Set-PnPField -List $listName -Identity $col.Name -Values @{Indexed = $true} -ErrorAction SilentlyContinue
            }

            Write-Host "    Added column: $($col.Name)" -ForegroundColor Gray
        }
    } else {
        Write-Host "    Column exists: $($col.Name)" -ForegroundColor DarkGray
    }
}

Write-Host "  Completed: $listName`n" -ForegroundColor Green

# ============================================================================
# LIST 3: JML_Policy_ApprovalDecisions
# Stores individual approval decisions
# ============================================================================

$listName = "JML_Policy_ApprovalDecisions"
Write-Host "Creating list: $listName..." -ForegroundColor White

$existingList = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($existingList) {
    Write-Host "  List already exists, skipping creation" -ForegroundColor Yellow
} else {
    if (-not $WhatIf) {
        New-PnPList -Title $listName -Template GenericList -EnableVersioning
        Write-Host "  List created" -ForegroundColor Green
    } else {
        Write-Host "  [WhatIf] Would create list" -ForegroundColor Gray
    }
}

# Add columns to ApprovalDecisions
$decisionColumns = @(
    @{ Name = "WorkflowId"; Type = "Number"; Required = $true; Indexed = $true },
    @{ Name = "PolicyId"; Type = "Number"; Required = $true; Indexed = $true },
    @{ Name = "StageNumber"; Type = "Number"; Required = $true },
    @{ Name = "StageName"; Type = "Text" },
    @{ Name = "ApproverId"; Type = "Number"; Required = $true; Indexed = $true },
    @{ Name = "ApproverName"; Type = "Text" },
    @{ Name = "ApproverEmail"; Type = "Text" },
    @{ Name = "OriginalApproverId"; Type = "Number" },  # For delegation tracking
    @{ Name = "OriginalApproverName"; Type = "Text" },
    @{ Name = "DelegatedById"; Type = "Number" },  # Who delegated
    @{ Name = "DelegatedByName"; Type = "Text" },
    @{ Name = "Status"; Type = "Choice"; Choices = @("Pending", "Approved", "Rejected", "Skipped", "Delegated", "Escalated", "TimedOut") },
    @{ Name = "Decision"; Type = "Choice"; Choices = @("Approved", "Rejected") },
    @{ Name = "Comments"; Type = "Note" },
    @{ Name = "RequestedDate"; Type = "DateTime"; Indexed = $true },
    @{ Name = "DueDate"; Type = "DateTime"; Indexed = $true },
    @{ Name = "DecisionDate"; Type = "DateTime" },
    @{ Name = "ReminderSentDate"; Type = "DateTime" },
    @{ Name = "EscalatedDate"; Type = "DateTime" },
    @{ Name = "IsRequired"; Type = "Boolean" },
    @{ Name = "CanDelegate"; Type = "Boolean" },
    @{ Name = "NotificationSent"; Type = "Boolean" },
    @{ Name = "TeamsCardSent"; Type = "Boolean" }
)

foreach ($col in $decisionColumns) {
    $existingField = Get-PnPField -List $listName -Identity $col.Name -ErrorAction SilentlyContinue
    if (-not $existingField) {
        if (-not $WhatIf) {
            switch ($col.Type) {
                "Text" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Text -AddToDefaultView
                }
                "Note" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Note
                }
                "Choice" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Choice -Choices $col.Choices -AddToDefaultView
                }
                "Boolean" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Boolean
                }
                "Number" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Number -AddToDefaultView
                }
                "DateTime" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type DateTime
                }
            }

            # Add index if specified
            if ($col.Indexed -and $col.Type -ne "Note") {
                Set-PnPField -List $listName -Identity $col.Name -Values @{Indexed = $true} -ErrorAction SilentlyContinue
            }

            Write-Host "    Added column: $($col.Name)" -ForegroundColor Gray
        }
    } else {
        Write-Host "    Column exists: $($col.Name)" -ForegroundColor DarkGray
    }
}

Write-Host "  Completed: $listName`n" -ForegroundColor Green

# ============================================================================
# LIST 4: JML_Policy_ApprovalDelegations
# Stores delegation assignments
# ============================================================================

$listName = "JML_Policy_ApprovalDelegations"
Write-Host "Creating list: $listName..." -ForegroundColor White

$existingList = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($existingList) {
    Write-Host "  List already exists, skipping creation" -ForegroundColor Yellow
} else {
    if (-not $WhatIf) {
        New-PnPList -Title $listName -Template GenericList -EnableVersioning
        Write-Host "  List created" -ForegroundColor Green
    } else {
        Write-Host "  [WhatIf] Would create list" -ForegroundColor Gray
    }
}

# Add columns to ApprovalDelegations
$delegationColumns = @(
    @{ Name = "DelegatorId"; Type = "Number"; Required = $true; Indexed = $true },
    @{ Name = "DelegatorName"; Type = "Text" },
    @{ Name = "DelegatorEmail"; Type = "Text" },
    @{ Name = "DelegateId"; Type = "Number"; Required = $true; Indexed = $true },
    @{ Name = "DelegateName"; Type = "Text" },
    @{ Name = "DelegateEmail"; Type = "Text" },
    @{ Name = "DelegationType"; Type = "Choice"; Choices = @("Temporary", "Permanent", "OutOfOffice") },
    @{ Name = "Scope"; Type = "Choice"; Choices = @("All", "SpecificPolicy", "Category") },
    @{ Name = "ScopePolicyId"; Type = "Number" },  # If Scope = SpecificPolicy
    @{ Name = "ScopeCategory"; Type = "Text" },  # If Scope = Category
    @{ Name = "StartDate"; Type = "DateTime"; Required = $true; Indexed = $true },
    @{ Name = "EndDate"; Type = "DateTime"; Indexed = $true },
    @{ Name = "Reason"; Type = "Note" },
    @{ Name = "IsActive"; Type = "Boolean"; Indexed = $true },
    @{ Name = "RevokedDate"; Type = "DateTime" },
    @{ Name = "RevokedById"; Type = "Number" },
    @{ Name = "RevokedByName"; Type = "Text" },
    @{ Name = "NotifyDelegator"; Type = "Boolean" },
    @{ Name = "NotifyDelegate"; Type = "Boolean" }
)

foreach ($col in $delegationColumns) {
    $existingField = Get-PnPField -List $listName -Identity $col.Name -ErrorAction SilentlyContinue
    if (-not $existingField) {
        if (-not $WhatIf) {
            switch ($col.Type) {
                "Text" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Text -AddToDefaultView
                }
                "Note" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Note
                }
                "Choice" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Choice -Choices $col.Choices -AddToDefaultView
                }
                "Boolean" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Boolean -AddToDefaultView
                }
                "Number" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Number
                }
                "DateTime" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type DateTime -AddToDefaultView
                }
            }

            # Add index if specified
            if ($col.Indexed -and $col.Type -ne "Note") {
                Set-PnPField -List $listName -Identity $col.Name -Values @{Indexed = $true} -ErrorAction SilentlyContinue
            }

            Write-Host "    Added column: $($col.Name)" -ForegroundColor Gray
        }
    } else {
        Write-Host "    Column exists: $($col.Name)" -ForegroundColor DarkGray
    }
}

Write-Host "  Completed: $listName`n" -ForegroundColor Green

# ============================================================================
# LIST 5: JML_Policy_EscalationRules
# Stores escalation rule configurations
# ============================================================================

$listName = "JML_Policy_EscalationRules"
Write-Host "Creating list: $listName..." -ForegroundColor White

$existingList = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($existingList) {
    Write-Host "  List already exists, skipping creation" -ForegroundColor Yellow
} else {
    if (-not $WhatIf) {
        New-PnPList -Title $listName -Template GenericList -EnableVersioning
        Write-Host "  List created" -ForegroundColor Green
    } else {
        Write-Host "  [WhatIf] Would create list" -ForegroundColor Gray
    }
}

# Add columns to EscalationRules
$escalationColumns = @(
    @{ Name = "RuleName"; Type = "Text"; Required = $true },
    @{ Name = "Description"; Type = "Note" },
    @{ Name = "TriggerType"; Type = "Choice"; Choices = @("HoursOverdue", "DaysOverdue", "StageTimeout", "NoResponse") },
    @{ Name = "TriggerValue"; Type = "Number"; Required = $true },  # Hours/Days threshold
    @{ Name = "ActionType"; Type = "Choice"; Choices = @("Notify", "NotifyManager", "AutoApprove", "AutoReject", "Reassign"); Required = $true },
    @{ Name = "ActionTargetId"; Type = "Number" },  # User ID for Reassign
    @{ Name = "ActionTargetName"; Type = "Text" },
    @{ Name = "ActionTargetEmail"; Type = "Text" },
    @{ Name = "NotifyOriginalApprover"; Type = "Boolean" },
    @{ Name = "NotifyInitiator"; Type = "Boolean" },
    @{ Name = "NotifyPolicyOwner"; Type = "Boolean" },
    @{ Name = "CustomEmailSubject"; Type = "Text" },
    @{ Name = "CustomEmailBody"; Type = "Note" },
    @{ Name = "AppliesTo"; Type = "Choice"; Choices = @("All", "Category", "Priority", "Template") },
    @{ Name = "AppliesToValue"; Type = "Text" },  # Category name, Priority level, or Template ID
    @{ Name = "IsActive"; Type = "Boolean"; Indexed = $true },
    @{ Name = "Priority"; Type = "Number" },  # Rule evaluation order
    @{ Name = "MaxEscalations"; Type = "Number" },  # Max times this rule can fire per workflow
    @{ Name = "CreatedById"; Type = "Number" },
    @{ Name = "ModifiedById"; Type = "Number" }
)

foreach ($col in $escalationColumns) {
    $existingField = Get-PnPField -List $listName -Identity $col.Name -ErrorAction SilentlyContinue
    if (-not $existingField) {
        if (-not $WhatIf) {
            switch ($col.Type) {
                "Text" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Text -AddToDefaultView
                }
                "Note" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Note
                }
                "Choice" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Choice -Choices $col.Choices -AddToDefaultView
                }
                "Boolean" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Boolean -AddToDefaultView
                }
                "Number" {
                    Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type Number
                }
            }

            # Add index if specified
            if ($col.Indexed) {
                Set-PnPField -List $listName -Identity $col.Name -Values @{Indexed = $true} -ErrorAction SilentlyContinue
            }

            Write-Host "    Added column: $($col.Name)" -ForegroundColor Gray
        }
    } else {
        Write-Host "    Column exists: $($col.Name)" -ForegroundColor DarkGray
    }
}

Write-Host "  Completed: $listName`n" -ForegroundColor Green

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Provisioning Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

Write-Host "`nLists created:" -ForegroundColor White
Write-Host "  1. JML_Policy_ApprovalTemplates   - Workflow templates" -ForegroundColor Gray
Write-Host "  2. JML_Policy_ApprovalWorkflows   - Active workflow instances" -ForegroundColor Gray
Write-Host "  3. JML_Policy_ApprovalDecisions   - Individual decisions" -ForegroundColor Gray
Write-Host "  4. JML_Policy_ApprovalDelegations - Delegation assignments" -ForegroundColor Gray
Write-Host "  5. JML_Policy_EscalationRules     - Escalation configurations" -ForegroundColor Gray

Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "  1. Configure default escalation rules" -ForegroundColor Gray
Write-Host "  2. Create workflow templates for common approval patterns" -ForegroundColor Gray
Write-Host "  3. Set up Teams webhook URL for adaptive card notifications" -ForegroundColor Gray
Write-Host "  4. Test approval workflow with sample policy" -ForegroundColor Gray
