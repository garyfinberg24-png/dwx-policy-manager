# ============================================================================
# DWx Policy Manager — Provision PM_PolicyDocuments Library
# ============================================================================
# Prerequisites:
#   Install-Module -Name PnP.PowerShell -Scope CurrentUser
#
# Usage:
#   Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive
#   .\provision-document-library.ps1
# ============================================================================

$ErrorActionPreference = "Stop"

# Check if library already exists
$existing = Get-PnPList -Identity "PM_PolicyDocuments" -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "  PM_PolicyDocuments already exists — adding any missing fields." -ForegroundColor Yellow
} else {
    Write-Host "  Creating PM_PolicyDocuments document library..." -ForegroundColor Green
    New-PnPList -Title "PM_PolicyDocuments" -Template DocumentLibrary -Url "PM_PolicyDocuments" | Out-Null
    Write-Host "  Library created." -ForegroundColor Green
}

# Define custom fields
$fields = @(
    @{ Name="PolicyId";               Type="Number";   DisplayName="Policy Id";               Required=$true }
    @{ Name="DocumentType";           Type="Choice";   DisplayName="Document Type";            Choices=@("Primary","Appendix","Form","Template","Guide","Reference"); Default="Primary" }
    @{ Name="DocumentCategory";       Type="Text";     DisplayName="Document Category" }
    @{ Name="DocumentTitle";          Type="Text";     DisplayName="Document Title" }
    @{ Name="DocumentDescription";    Type="Note";     DisplayName="Document Description" }
    @{ Name="DocumentVersion";        Type="Text";     DisplayName="Document Version";         Default="1.0" }
    @{ Name="DocumentVersionDate";    Type="DateTime"; DisplayName="Version Date" }
    @{ Name="IsCurrentVersion";       Type="Boolean";  DisplayName="Is Current Version";       Default="1" }
    @{ Name="SecurityClassification"; Type="Choice";   DisplayName="Security Classification";  Choices=@("Public","Internal","Confidential","Restricted"); Default="Internal" }
    @{ Name="RequiresApproval";       Type="Boolean";  DisplayName="Requires Approval";        Default="0" }
    @{ Name="RestrictedAccess";       Type="Boolean";  DisplayName="Restricted Access";        Default="0" }
    @{ Name="ViewCount";              Type="Number";   DisplayName="View Count" }
    @{ Name="DownloadCount";          Type="Number";   DisplayName="Download Count" }
    @{ Name="LastViewedDate";         Type="DateTime"; DisplayName="Last Viewed Date" }
    @{ Name="IsActive";               Type="Boolean";  DisplayName="Is Active";                Default="1" }
    @{ Name="IsArchived";             Type="Boolean";  DisplayName="Is Archived";              Default="0" }
    @{ Name="IsFeatured";             Type="Boolean";  DisplayName="Is Featured";              Default="0" }
    @{ Name="IsPopular";              Type="Boolean";  DisplayName="Is Popular";               Default="0" }
    @{ Name="SearchKeywords";         Type="Note";     DisplayName="Search Keywords" }
    @{ Name="Tags";                   Type="Note";     DisplayName="Tags (JSON)" }
)

Write-Host "`n  Adding fields..." -ForegroundColor Cyan
foreach ($field in $fields) {
    $existingField = Get-PnPField -List "PM_PolicyDocuments" -Identity $field.Name -ErrorAction SilentlyContinue
    if ($existingField) {
        Write-Host "    [EXISTS] $($field.Name)" -ForegroundColor DarkGray
        continue
    }

    $params = @{
        List         = "PM_PolicyDocuments"
        InternalName = $field.Name
        DisplayName  = $field.DisplayName
        Type         = $field.Type
        Group        = "PM Policy Manager"
        Required     = $field.Required -eq $true
    }
    if ($field.Choices) { $params.Choices = $field.Choices }

    Add-PnPField @params | Out-Null

    if ($field.Default) {
        Set-PnPField -List "PM_PolicyDocuments" -Identity $field.Name -Values @{DefaultValue = $field.Default }
    }

    Write-Host "    [ADDED]  $($field.Name) ($($field.Type))" -ForegroundColor Green
}

# ============================================================================
# DONE
# ============================================================================
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  PM_PolicyDocuments provisioned!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Next steps:" -ForegroundColor Yellow
Write-Host "    1. Upload policy PDFs to PM_PolicyDocuments"
Write-Host "    2. For each uploaded file, set the PolicyId field to match"
Write-Host "       the corresponding PM_Policies list item Id"
Write-Host "    3. Update the DocumentURL field on PM_Policies records:"
Write-Host "       /sites/PolicyManager/PM_PolicyDocuments/YourPolicy.pdf"
Write-Host ""
