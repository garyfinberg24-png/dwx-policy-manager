# =============================================================================
# 26-SLABreaches-List.ps1
# Creates PM_SLABreaches list for persisting SLA breach records
# Assumes: Already connected to SharePoint via Connect-PnPOnline
# Site: https://mf7m.sharepoint.com/sites/PolicyManager
# =============================================================================

$listName = "PM_SLABreaches"

# Check if list already exists
$existing = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "[SKIP] $listName already exists" -ForegroundColor Yellow
} else {
    Write-Host "Creating $listName..." -ForegroundColor Cyan
    New-PnPList -Title $listName -Template GenericList -Url "Lists/$listName"
    Write-Host "[OK] $listName created" -ForegroundColor Green
}

# Add columns (idempotent — checks before creating)
$columns = @(
    @{ Name = "PolicyId";          Type = "Number";  },
    @{ Name = "PolicyTitle";       Type = "Text";    },
    @{ Name = "PolicyNumber";      Type = "Text";    },
    @{ Name = "SLAType";           Type = "Choice";  Choices = @("Acknowledgement", "Approval", "Review", "Authoring") },
    @{ Name = "TargetDays";        Type = "Number";  },
    @{ Name = "ActualDays";        Type = "Number";  },
    @{ Name = "DaysOverdue";       Type = "Number";  },
    @{ Name = "BreachedDate";      Type = "DateTime"; },
    @{ Name = "DetectedDate";      Type = "DateTime"; },
    @{ Name = "ResponsibleUserId"; Type = "Number";  },
    @{ Name = "ResponsibleEmail";  Type = "Text";    },
    @{ Name = "ResponsibleName";   Type = "Text";    },
    @{ Name = "BreachStatus";      Type = "Choice";  Choices = @("Open", "Acknowledged", "Resolved", "Waived") },
    @{ Name = "ResolvedDate";      Type = "DateTime"; },
    @{ Name = "ResolvedBy";        Type = "Text";    },
    @{ Name = "Resolution";        Type = "Note";    },
    @{ Name = "Severity";          Type = "Choice";  Choices = @("Critical", "High", "Medium", "Low") },
    @{ Name = "EntityId";          Type = "Number";  },
    @{ Name = "EntityType";        Type = "Text";    },
    @{ Name = "ComplianceRelevant"; Type = "Boolean"; }
)

foreach ($col in $columns) {
    $existingField = Get-PnPField -List $listName -Identity $col.Name -ErrorAction SilentlyContinue
    if ($existingField) {
        Write-Host "  [SKIP] $($col.Name) already exists" -ForegroundColor Yellow
        continue
    }

    $params = @{
        List         = $listName
        InternalName = $col.Name
        DisplayName  = $col.Name
        Type         = $col.Type
        Group        = "Policy Manager"
    }

    if ($col.Type -eq "Choice" -and $col.Choices) {
        $params["Choices"] = $col.Choices
    }

    Add-PnPField @params | Out-Null
    Write-Host "  [OK] $($col.Name) ($($col.Type))" -ForegroundColor Green
}

# Set default values
Set-PnPField -List $listName -Identity "BreachStatus" -Values @{ DefaultValue = "Open" }
Set-PnPField -List $listName -Identity "ComplianceRelevant" -Values @{ DefaultValue = "1" }

# Create indexes for common queries
Add-PnPFieldToView -List $listName -View "All Items" -Field "PolicyTitle"
Add-PnPFieldToView -List $listName -View "All Items" -Field "SLAType"
Add-PnPFieldToView -List $listName -View "All Items" -Field "BreachStatus"
Add-PnPFieldToView -List $listName -View "All Items" -Field "DaysOverdue"
Add-PnPFieldToView -List $listName -View "All Items" -Field "Severity"
Add-PnPFieldToView -List $listName -View "All Items" -Field "DetectedDate"

Write-Host "`n[DONE] $listName provisioned with $($columns.Count) columns" -ForegroundColor Green
