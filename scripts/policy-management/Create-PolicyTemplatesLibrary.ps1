# Create Policy Templates and Source Documents Libraries
# This script creates the infrastructure for policy authoring enhancements

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl
)

# Connect to SharePoint
Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Host "Creating Policy Templates and Source Documents infrastructure..." -ForegroundColor Cyan

# ============================================================================
# 1. Policy Templates List (Approved Templates for Authors)
# ============================================================================
Write-Host "Creating JML_PolicyTemplates list..." -ForegroundColor Yellow

$templatesList = Get-PnPList -Identity "JML_PolicyTemplates" -ErrorAction SilentlyContinue
if ($null -eq $templatesList) {
    New-PnPList -Title "JML_PolicyTemplates" -Template GenericList -OnQuickLaunch
    Write-Host "  ✓ List created" -ForegroundColor Green
}

# Add fields to Policy Templates
Add-PnPField -List "JML_PolicyTemplates" -DisplayName "TemplateType" -InternalName "TemplateType" -Type Choice -Choices "Standard Policy","Procedure","Guideline","Code of Conduct","Health & Safety","IT Policy","HR Policy","Financial Policy","Legal Policy","Data Privacy","Security Policy","Custom" -AddToDefaultView

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "TemplateCategory" -InternalName "TemplateCategory" -Type Choice -Choices "General","HR","IT","Finance","Legal","Operations","Compliance","Health & Safety","Data Privacy","Security","Quality" -AddToDefaultView

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "TemplateDescription" -InternalName "TemplateDescription" -Type Note

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "TemplateContent" -InternalName "TemplateContent" -Type Note

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -AddToDefaultView
Set-PnPField -List "JML_PolicyTemplates" -Identity "IsActive" -Values @{DefaultValue="1"}

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "UsageCount" -InternalName "UsageCount" -Type Number -AddToDefaultView
Set-PnPField -List "JML_PolicyTemplates" -Identity "UsageCount" -Values @{DefaultValue="0"}

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "ComplianceRisk" -InternalName "ComplianceRisk" -Type Choice -Choices "High","Medium","Low" -AddToDefaultView

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "SuggestedReadTimeframe" -InternalName "SuggestedReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6" -AddToDefaultView

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "RequiresAcknowledgement" -InternalName "RequiresAcknowledgement" -Type Boolean
Set-PnPField -List "JML_PolicyTemplates" -Identity "RequiresAcknowledgement" -Values @{DefaultValue="1"}

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "RequiresQuiz" -InternalName "RequiresQuiz" -Type Boolean
Set-PnPField -List "JML_PolicyTemplates" -Identity "RequiresQuiz" -Values @{DefaultValue="0"}

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "KeyPointsTemplate" -InternalName "KeyPointsTemplate" -Type Note

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "ApprovedBy" -InternalName "ApprovedBy" -Type User

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "ApprovedDate" -InternalName "ApprovedDate" -Type DateTime

Add-PnPField -List "JML_PolicyTemplates" -DisplayName "Tags" -InternalName "Tags" -Type Note

Write-Host "  ✓ Policy Templates list configured with 15 fields" -ForegroundColor Green

# ============================================================================
# 2. Policy Source Documents Library (Uploaded Files)
# ============================================================================
Write-Host "Creating JML_PolicySourceDocuments library..." -ForegroundColor Yellow

$sourceDocsLib = Get-PnPList -Identity "JML_PolicySourceDocuments" -ErrorAction SilentlyContinue
if ($null -eq $sourceDocsLib) {
    New-PnPList -Title "JML_PolicySourceDocuments" -Template DocumentLibrary -OnQuickLaunch
    Write-Host "  ✓ Library created" -ForegroundColor Green
}

# Add metadata fields
Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -AddToDefaultView

Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "DocumentType" -InternalName "DocumentType" -Type Choice -Choices "Word Document","Excel Spreadsheet","PowerPoint Presentation","PDF","Image","Video","Other" -AddToDefaultView

Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "PolicyCategory" -InternalName "PolicyCategory" -Type Choice -Choices "General","HR","IT","Finance","Legal","Operations","Compliance","Health & Safety","Data Privacy","Security","Quality" -AddToDefaultView

Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "UploadedBy" -InternalName "UploadedBy" -Type User -AddToDefaultView

Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "UploadDate" -InternalName "UploadDate" -Type DateTime -AddToDefaultView

Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "FileStatus" -InternalName "FileStatus" -Type Choice -Choices "Uploaded","Processing","Converted","Published","Archived" -AddToDefaultView

Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "ExtractedText" -InternalName "ExtractedText" -Type Note

Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "PageCount" -InternalName "PageCount" -Type Number

Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "WordCount" -InternalName "WordCount" -Type Number

Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "FileHash" -InternalName "FileHash" -Type Text

Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "VirusScanStatus" -InternalName "VirusScanStatus" -Type Choice -Choices "Pending","Clean","Infected","Error"
Set-PnPField -List "JML_PolicySourceDocuments" -Identity "VirusScanStatus" -Values @{DefaultValue="Clean"}

Add-PnPField -List "JML_PolicySourceDocuments" -DisplayName "ProcessingNotes" -InternalName "ProcessingNotes" -Type Note

Write-Host "  ✓ Policy Source Documents library configured with 12 metadata fields" -ForegroundColor Green

# ============================================================================
# 3. Policy Reviewers List (Reviewers and Approvers)
# ============================================================================
Write-Host "Creating JML_PolicyReviewers list..." -ForegroundColor Yellow

$reviewersList = Get-PnPList -Identity "JML_PolicyReviewers" -ErrorAction SilentlyContinue
if ($null -eq $reviewersList) {
    New-PnPList -Title "JML_PolicyReviewers" -Template GenericList -OnQuickLaunch
    Write-Host "  ✓ List created" -ForegroundColor Green
}

Add-PnPField -List "JML_PolicyReviewers" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -AddToDefaultView

Add-PnPField -List "JML_PolicyReviewers" -DisplayName "Reviewer" -InternalName "Reviewer" -Type User -AddToDefaultView

Add-PnPField -List "JML_PolicyReviewers" -DisplayName "ReviewerType" -InternalName "ReviewerType" -Type Choice -Choices "Technical Reviewer","Legal Reviewer","Compliance Reviewer","Department Head","Executive Approver","Final Approver" -AddToDefaultView

Add-PnPField -List "JML_PolicyReviewers" -DisplayName "ReviewStatus" -InternalName "ReviewStatus" -Type Choice -Choices "Pending","In Review","Approved","Rejected","Revision Requested" -AddToDefaultView
Set-PnPField -List "JML_PolicyReviewers" -Identity "ReviewStatus" -Values @{DefaultValue="Pending"}

Add-PnPField -List "JML_PolicyReviewers" -DisplayName "AssignedDate" -InternalName "AssignedDate" -Type DateTime -AddToDefaultView

Add-PnPField -List "JML_PolicyReviewers" -DisplayName "ReviewedDate" -InternalName "ReviewedDate" -Type DateTime

Add-PnPField -List "JML_PolicyReviewers" -DisplayName "ReviewComments" -InternalName "ReviewComments" -Type Note

Add-PnPField -List "JML_PolicyReviewers" -DisplayName "ReviewSequence" -InternalName "ReviewSequence" -Type Number -AddToDefaultView

Add-PnPField -List "JML_PolicyReviewers" -DisplayName "IsMandatory" -InternalName "IsMandatory" -Type Boolean -AddToDefaultView
Set-PnPField -List "JML_PolicyReviewers" -Identity "IsMandatory" -Values @{DefaultValue="1"}

Add-PnPField -List "JML_PolicyReviewers" -DisplayName "DueDays" -InternalName "DueDays" -Type Number
Set-PnPField -List "JML_PolicyReviewers" -Identity "DueDays" -Values @{DefaultValue="5"}

Add-PnPField -List "JML_PolicyReviewers" -DisplayName "NotificationSent" -InternalName "NotificationSent" -Type Boolean

Write-Host "  ✓ Policy Reviewers list configured with 11 fields" -ForegroundColor Green

# ============================================================================
# 4. Policy Metadata Profiles (Pre-filled metadata sets)
# ============================================================================
Write-Host "Creating JML_PolicyMetadataProfiles list..." -ForegroundColor Yellow

$metadataProfilesList = Get-PnPList -Identity "JML_PolicyMetadataProfiles" -ErrorAction SilentlyContinue
if ($null -eq $metadataProfilesList) {
    New-PnPList -Title "JML_PolicyMetadataProfiles" -Template GenericList -OnQuickLaunch
    Write-Host "  ✓ List created" -ForegroundColor Green
}

Add-PnPField -List "JML_PolicyMetadataProfiles" -DisplayName "ProfileName" -InternalName "ProfileName" -Type Text -AddToDefaultView

Add-PnPField -List "JML_PolicyMetadataProfiles" -DisplayName "PolicyCategory" -InternalName "PolicyCategory" -Type Choice -Choices "General","HR","IT","Finance","Legal","Operations","Compliance","Health & Safety","Data Privacy","Security","Quality" -AddToDefaultView

Add-PnPField -List "JML_PolicyMetadataProfiles" -DisplayName "ComplianceRisk" -InternalName "ComplianceRisk" -Type Choice -Choices "High","Medium","Low" -AddToDefaultView

Add-PnPField -List "JML_PolicyMetadataProfiles" -DisplayName "ReadTimeframe" -InternalName "ReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6"

Add-PnPField -List "JML_PolicyMetadataProfiles" -DisplayName "RequiresAcknowledgement" -InternalName "RequiresAcknowledgement" -Type Boolean
Set-PnPField -List "JML_PolicyMetadataProfiles" -Identity "RequiresAcknowledgement" -Values @{DefaultValue="1"}

Add-PnPField -List "JML_PolicyMetadataProfiles" -DisplayName "RequiresQuiz" -InternalName "RequiresQuiz" -Type Boolean

Add-PnPField -List "JML_PolicyMetadataProfiles" -DisplayName "DefaultReviewers" -InternalName "DefaultReviewers" -Type UserMulti

Add-PnPField -List "JML_PolicyMetadataProfiles" -DisplayName "DefaultApprovers" -InternalName "DefaultApprovers" -Type UserMulti

Add-PnPField -List "JML_PolicyMetadataProfiles" -DisplayName "TargetDepartments" -InternalName "TargetDepartments" -Type Note

Add-PnPField -List "JML_PolicyMetadataProfiles" -DisplayName "TargetRoles" -InternalName "TargetRoles" -Type Note

Add-PnPField -List "JML_PolicyMetadataProfiles" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean
Set-PnPField -List "JML_PolicyMetadataProfiles" -Identity "IsActive" -Values @{DefaultValue="1"}

Write-Host "  ✓ Policy Metadata Profiles list configured with 11 fields" -ForegroundColor Green

# ============================================================================
# 5. File Conversion Queue (Track file processing)
# ============================================================================
Write-Host "Creating JML_FileConversionQueue list..." -ForegroundColor Yellow

$conversionQueueList = Get-PnPList -Identity "JML_FileConversionQueue" -ErrorAction SilentlyContinue
if ($null -eq $conversionQueueList) {
    New-PnPList -Title "JML_FileConversionQueue" -Template GenericList -OnQuickLaunch
    Write-Host "  ✓ List created" -ForegroundColor Green
}

Add-PnPField -List "JML_FileConversionQueue" -DisplayName "SourceFileUrl" -InternalName "SourceFileUrl" -Type URL -AddToDefaultView

Add-PnPField -List "JML_FileConversionQueue" -DisplayName "SourceFileType" -InternalName "SourceFileType" -Type Text

Add-PnPField -List "JML_FileConversionQueue" -DisplayName "QueueStatus" -InternalName "QueueStatus" -Type Choice -Choices "Queued","Processing","Completed","Failed" -AddToDefaultView
Set-PnPField -List "JML_FileConversionQueue" -Identity "QueueStatus" -Values @{DefaultValue="Queued"}

Add-PnPField -List "JML_FileConversionQueue" -DisplayName "QueuedDate" -InternalName "QueuedDate" -Type DateTime -AddToDefaultView

Add-PnPField -List "JML_FileConversionQueue" -DisplayName "ProcessedDate" -InternalName "ProcessedDate" -Type DateTime

Add-PnPField -List "JML_FileConversionQueue" -DisplayName "ConvertedContent" -InternalName "ConvertedContent" -Type Note

Add-PnPField -List "JML_FileConversionQueue" -DisplayName "ErrorMessage" -InternalName "ErrorMessage" -Type Note

Add-PnPField -List "JML_FileConversionQueue" -DisplayName "ProcessingTime" -InternalName "ProcessingTime" -Type Number

Add-PnPField -List "JML_FileConversionQueue" -DisplayName "SubmittedBy" -InternalName "SubmittedBy" -Type User

Write-Host "  ✓ File Conversion Queue list configured with 9 fields" -ForegroundColor Green

# ============================================================================
# Summary
# ============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Policy Authoring Infrastructure Created" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "✓ JML_PolicyTemplates (15 fields)" -ForegroundColor Green
Write-Host "✓ JML_PolicySourceDocuments (12 fields)" -ForegroundColor Green
Write-Host "✓ JML_PolicyReviewers (11 fields)" -ForegroundColor Green
Write-Host "✓ JML_PolicyMetadataProfiles (11 fields)" -ForegroundColor Green
Write-Host "✓ JML_FileConversionQueue (9 fields)" -ForegroundColor Green
Write-Host ""
Write-Host "Next: Run Add-PolicyTemplateSampleData.ps1 to populate with sample data" -ForegroundColor Yellow
Write-Host ""
