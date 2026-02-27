# ============================================================================
# Policy Manager - Provision PM_PolicyRequests List
# Creates the list for storing policy requests from the Request Policy wizard
# Target: https://mf7m.sharepoint.com/sites/PolicyManager
# ============================================================================
#
# USAGE:
#   Connect-PnPOnline first, then:
#   .\Create-PM_PolicyRequests.ps1
#
# PREREQUISITES:
#   - PnP.PowerShell module installed
#   - Already connected to https://mf7m.sharepoint.com/sites/PolicyManager
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Policy Manager - Provision PM_PolicyRequests" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ============================================================================
# Create PM_PolicyRequests list
# ============================================================================

$listName = "PM_PolicyRequests"

Write-Host "Creating list: $listName" -ForegroundColor Yellow
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

# ============================================================================
# Add Fields
# ============================================================================

Write-Host ""
Write-Host "Adding fields to $listName..." -ForegroundColor Yellow

$fields = @(
    @{Name="RequestedBy"; Type="Text"; Required=$true},
    @{Name="RequestedByEmail"; Type="Text"; Required=$false},
    @{Name="RequestedByDepartment"; Type="Text"; Required=$false},
    @{Name="PolicyCategory"; Type="Choice"; Choices="HR Policies,IT & Security,Health & Safety,Compliance,Financial,Operational,Legal,Environmental,Quality Assurance,Data Privacy,Custom"; Required=$true},
    @{Name="PolicyType"; Type="Choice"; Choices="New Policy,Policy Update,Policy Review,Policy Replacement"; Required=$false},
    @{Name="Priority"; Type="Choice"; Choices="Low,Medium,High,Critical"; Required=$false},
    @{Name="TargetAudience"; Type="Text"; Required=$false},
    @{Name="BusinessJustification"; Type="Note"; Required=$true},
    @{Name="RegulatoryDriver"; Type="Text"; Required=$false},
    @{Name="DesiredEffectiveDate"; Type="DateTime"; Required=$false},
    @{Name="ReadTimeframeDays"; Type="Number"; Required=$false},
    @{Name="RequiresAcknowledgement"; Type="Boolean"; Required=$false},
    @{Name="RequiresQuiz"; Type="Boolean"; Required=$false},
    @{Name="AdditionalNotes"; Type="Note"; Required=$false},
    @{Name="NotifyAuthors"; Type="Boolean"; Required=$false},
    @{Name="PreferredAuthor"; Type="Text"; Required=$false},
    @{Name="AttachmentUrls"; Type="Note"; Required=$false},
    @{Name="Status"; Type="Choice"; Choices="New,Assigned,InProgress,Draft Ready,Completed,Rejected"; Required=$false},
    @{Name="AssignedAuthor"; Type="Text"; Required=$false},
    @{Name="AssignedAuthorEmail"; Type="Text"; Required=$false},
    @{Name="ReferenceNumber"; Type="Text"; Required=$false}
)

$added = 0
$skipped = 0

foreach ($field in $fields) {
    $existingField = Get-PnPField -List $listName -Identity $field.Name -ErrorAction SilentlyContinue
    if ($null -ne $existingField) {
        Write-Host "    Exists: $($field.Name)" -ForegroundColor Gray
        $skipped++
        continue
    }

    $params = @{
        List = $listName
        DisplayName = $field.Name
        InternalName = $field.Name
        Type = $field.Type
        Required = ($field.Required -eq $true)
        AddToDefaultView = $true
    }

    if ($field.Type -eq "Choice" -and $field.Choices) {
        $choiceArray = $field.Choices -split ","
        $params["Choices"] = $choiceArray
    }

    try {
        Add-PnPField @params | Out-Null
        Write-Host "    Added: $($field.Name) ($($field.Type))" -ForegroundColor Green
        $added++
    } catch {
        Write-Host "    FAILED: $($field.Name) - $_" -ForegroundColor Red
    }
}

# ============================================================================
# Summary
# ============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Provisioning Complete" -ForegroundColor Cyan
Write-Host "------------------------------------------------------------" -ForegroundColor Gray
Write-Host "  List: $listName" -ForegroundColor White
Write-Host "  Fields added: $added" -ForegroundColor Green
Write-Host "  Fields skipped: $skipped (already existed)" -ForegroundColor Gray
Write-Host "  Total fields: $($fields.Count)" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
