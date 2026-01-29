# ============================================================================
# Create PM_PolicyMetadataProfiles List
# Run this script to create the missing PolicyMetadataProfiles list
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"
)

# Connect to SharePoint Online
Write-Host "Connecting to SharePoint site: $SiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

# Check if list already exists
$listName = "PM_PolicyMetadataProfiles"
$existingList = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($existingList) {
    Write-Host "List '$listName' already exists. Skipping creation." -ForegroundColor Yellow
} else {
    Write-Host "Creating list: $listName" -ForegroundColor Green

    # Create the list
    $list = New-PnPList -Title $listName -Template GenericList -EnableVersioning

    # Add columns
    Write-Host "Adding columns to $listName..." -ForegroundColor Cyan

    # Profile Name
    Add-PnPField -List $listName -DisplayName "ProfileName" -InternalName "ProfileName" -Type Text -Required

    # Description
    Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note

    # Is Active
    Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean
    Set-PnPField -List $listName -Identity "IsActive" -Values @{DefaultValue = "1"}

    # Profile Type (e.g., HR, Legal, Compliance, IT, General)
    Add-PnPField -List $listName -DisplayName "ProfileType" -InternalName "ProfileType" -Type Choice -Choices "HR","Legal","Compliance","IT","Finance","Operations","General"

    # Required Fields (JSON array of field names)
    Add-PnPField -List $listName -DisplayName "RequiredFields" -InternalName "RequiredFields" -Type Note

    # Optional Fields (JSON array of field names)
    Add-PnPField -List $listName -DisplayName "OptionalFields" -InternalName "OptionalFields" -Type Note

    # Default Values (JSON object)
    Add-PnPField -List $listName -DisplayName "DefaultValues" -InternalName "DefaultValues" -Type Note

    # Validation Rules (JSON object)
    Add-PnPField -List $listName -DisplayName "ValidationRules" -InternalName "ValidationRules" -Type Note

    # Sort Order
    Add-PnPField -List $listName -DisplayName "SortOrder" -InternalName "SortOrder" -Type Number

    # Icon (for UI display)
    Add-PnPField -List $listName -DisplayName "Icon" -InternalName "Icon" -Type Text

    Write-Host "List '$listName' created successfully!" -ForegroundColor Green

    # Add sample data
    Write-Host "Adding sample metadata profiles..." -ForegroundColor Cyan

    $profiles = @(
        @{
            Title = "General Policy"
            ProfileName = "General Policy"
            Description = "Standard policy metadata profile for general organizational policies"
            IsActive = $true
            ProfileType = "General"
            RequiredFields = '["Title","Description","EffectiveDate","Owner"]'
            OptionalFields = '["ReviewDate","Category","Tags"]'
            DefaultValues = '{}'
            SortOrder = 1
            Icon = "Document"
        },
        @{
            Title = "HR Policy"
            ProfileName = "HR Policy"
            Description = "Metadata profile for Human Resources policies"
            IsActive = $true
            ProfileType = "HR"
            RequiredFields = '["Title","Description","EffectiveDate","Owner","Department","TargetAudience"]'
            OptionalFields = '["ReviewDate","ComplianceRequirement","TrainingRequired"]'
            DefaultValues = '{"Department":"Human Resources"}'
            SortOrder = 2
            Icon = "People"
        },
        @{
            Title = "Compliance Policy"
            ProfileName = "Compliance Policy"
            Description = "Metadata profile for regulatory and compliance policies"
            IsActive = $true
            ProfileType = "Compliance"
            RequiredFields = '["Title","Description","EffectiveDate","Owner","RegulatoryBody","ComplianceFramework"]'
            OptionalFields = '["AuditFrequency","PenaltyRisk","LastAuditDate"]'
            DefaultValues = '{"RequiresAcknowledgement":true}'
            SortOrder = 3
            Icon = "Shield"
        },
        @{
            Title = "IT Security Policy"
            ProfileName = "IT Security Policy"
            Description = "Metadata profile for IT and security policies"
            IsActive = $true
            ProfileType = "IT"
            RequiredFields = '["Title","Description","EffectiveDate","Owner","SecurityLevel","DataClassification"]'
            OptionalFields = '["TechnicalControls","IncidentResponse"]'
            DefaultValues = '{"RequiresQuiz":true}'
            SortOrder = 4
            Icon = "Lock"
        }
    )

    foreach ($profile in $profiles) {
        Add-PnPListItem -List $listName -Values $profile | Out-Null
        Write-Host "  Added profile: $($profile.ProfileName)" -ForegroundColor Gray
    }

    Write-Host "Sample data added successfully!" -ForegroundColor Green
}

Write-Host "`nScript completed!" -ForegroundColor Green
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. Update PolicyAuthorEnhanced.tsx to use PM_LISTS constants" -ForegroundColor White
Write-Host "  2. Rebuild and redeploy the SPFx solution" -ForegroundColor White
