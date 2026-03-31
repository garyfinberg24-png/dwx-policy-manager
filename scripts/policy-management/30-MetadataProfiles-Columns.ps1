# ============================================================================
# 30-MetadataProfiles-Columns.ps1
# Adds missing columns to PM_PolicyMetadataProfiles for full metadata support
# Assumes user is already connected to SharePoint
# ============================================================================

$listName = "PM_PolicyMetadataProfiles"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if (-not $list) {
    Write-Host "[SKIP] $listName does not exist. Run Create-PM_PolicyMetadataProfiles.ps1 first." -ForegroundColor Yellow
    return
}

Write-Host "Adding columns to $listName..." -ForegroundColor Cyan

# Core metadata fields (may already exist from AdminConfigService writes)
$fields = @(
    @{ Name="PolicyCategory";     Type="Text";    DisplayName="Policy Category" },
    @{ Name="ComplianceRisk";     Type="Choice";  DisplayName="Compliance Risk";  Choices=@("Critical","High","Medium","Low","Informational") },
    @{ Name="ReadTimeframe";      Type="Choice";  DisplayName="Read Timeframe";   Choices=@("Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6") },
    @{ Name="RequiresAcknowledgement"; Type="Boolean"; DisplayName="Requires Acknowledgement" },
    @{ Name="RequiresQuiz";       Type="Boolean"; DisplayName="Requires Quiz" },
    @{ Name="TargetDepartments";  Type="Note";    DisplayName="Target Departments" },
    @{ Name="DistributionScope";  Type="Choice";  DisplayName="Distribution Scope"; Choices=@("All Employees","Department Only","Role-Based","Security Group") },
    @{ Name="TemplateType";       Type="Choice";  DisplayName="Template Type";    Choices=@("word","excel","powerpoint","html","infographic") },
    @{ Name="DocumentTemplateId"; Type="Text";    DisplayName="Document Template ID" },
    @{ Name="Classification";     Type="Choice";  DisplayName="Classification";   Choices=@("Public","Internal","Confidential","Restricted") },
    @{ Name="RegulatoryFramework"; Type="Choice"; DisplayName="Regulatory Framework"; Choices=@("None","POPIA","GDPR","OHS","BCEA","FICA","KING_IV","ISO27001","ISO9001") },
    @{ Name="ReviewCycleMonths";  Type="Number";  DisplayName="Review Cycle (Months)" },
    @{ Name="EstimatedReadTimeMinutes"; Type="Number"; DisplayName="Estimated Read Time (Minutes)" },
    @{ Name="RetentionYears";     Type="Number";  DisplayName="Retention Period (Years)" },
    @{ Name="AutoNotifyOnUpdate"; Type="Boolean"; DisplayName="Auto-Notify on Update" },
    @{ Name="RequiresDigitalSignature"; Type="Boolean"; DisplayName="Requires Digital Signature" },
    @{ Name="TargetAudiences";    Type="Note";    DisplayName="Target Audiences" },
    @{ Name="TargetSecurityGroups"; Type="Note";  DisplayName="Target Security Groups" }
)

foreach ($field in $fields) {
    $existing = Get-PnPField -List $listName -Identity $field.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  [EXISTS] $($field.Name)" -ForegroundColor DarkGray
    } else {
        $params = @{
            List = $listName
            DisplayName = $field.DisplayName
            InternalName = $field.Name
            Type = $field.Type
            ErrorAction = "SilentlyContinue"
        }
        if ($field.Choices) {
            $params["Choices"] = $field.Choices
        }
        Add-PnPField @params
        Write-Host "  [OK] $($field.Name) ($($field.Type))" -ForegroundColor Green
    }
}

# Set defaults for boolean fields
try {
    Set-PnPField -List $listName -Identity "RequiresAcknowledgement" -Values @{DefaultValue = "1"} -ErrorAction SilentlyContinue
    Set-PnPField -List $listName -Identity "AutoNotifyOnUpdate" -Values @{DefaultValue = "1"} -ErrorAction SilentlyContinue
} catch { }

Write-Host "`n[DONE] $listName column patch complete." -ForegroundColor Green
