# ============================================================================
# PM_PolicySourceDocuments - Document Library Provisioning
# Creates the document library used by the Policy Builder to store
# Office documents (Word, Excel, PowerPoint) created via the wizard.
# ============================================================================
# Prerequisites: Already connected to SharePoint via Connect-PnPOnline
# Site: https://mf7m.sharepoint.com/sites/PolicyManager
# ============================================================================

$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"
$LibraryName = "PM_PolicySourceDocuments"
$LibraryTitle = "Policy Source Documents"

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " Creating $LibraryName Document Library" -ForegroundColor Cyan
Write-Host "============================================`n" -ForegroundColor Cyan

# Check if library already exists
$existingList = Get-PnPList -Identity $LibraryName -ErrorAction SilentlyContinue

if ($existingList) {
    Write-Host "[SKIP] $LibraryName already exists." -ForegroundColor Yellow
} else {
    Write-Host "[CREATE] Creating document library: $LibraryName..." -ForegroundColor Green
    New-PnPList -Title $LibraryTitle -Url $LibraryName -Template DocumentLibrary -EnableVersioning
    Write-Host "[OK] Document library created." -ForegroundColor Green
}

# Get the library
$list = Get-PnPList -Identity $LibraryName

# Add custom columns
$fields = @(
    @{ Name = "DocumentType"; Type = "Choice"; Choices = @("Word Document", "Excel Spreadsheet", "PowerPoint Presentation", "PDF", "Image", "Other"); Default = "Word Document" },
    @{ Name = "FileStatus"; Type = "Choice"; Choices = @("Draft", "Uploaded", "Processing", "Ready", "Archived"); Default = "Draft" },
    @{ Name = "PolicyTitle"; Type = "Text" },
    @{ Name = "PolicyId"; Type = "Number" },
    @{ Name = "UploadDate"; Type = "DateTime" },
    @{ Name = "CreatedByWizard"; Type = "Boolean"; Default = "1" }
)

foreach ($field in $fields) {
    $existing = Get-PnPField -List $LibraryName -Identity $field.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "[SKIP] Field '$($field.Name)' already exists." -ForegroundColor Yellow
        continue
    }

    switch ($field.Type) {
        "Choice" {
            Add-PnPField -List $LibraryName -DisplayName $field.Name -InternalName $field.Name -Type Choice -Choices $field.Choices -DefaultValue $field.Default -AddToDefaultView
            Write-Host "[OK] Added Choice field: $($field.Name)" -ForegroundColor Green
        }
        "Text" {
            Add-PnPField -List $LibraryName -DisplayName $field.Name -InternalName $field.Name -Type Text -AddToDefaultView
            Write-Host "[OK] Added Text field: $($field.Name)" -ForegroundColor Green
        }
        "Number" {
            Add-PnPField -List $LibraryName -DisplayName $field.Name -InternalName $field.Name -Type Number -AddToDefaultView
            Write-Host "[OK] Added Number field: $($field.Name)" -ForegroundColor Green
        }
        "DateTime" {
            Add-PnPField -List $LibraryName -DisplayName $field.Name -InternalName $field.Name -Type DateTime -AddToDefaultView
            Write-Host "[OK] Added DateTime field: $($field.Name)" -ForegroundColor Green
        }
        "Boolean" {
            Add-PnPField -List $LibraryName -DisplayName $field.Name -InternalName $field.Name -Type Boolean -AddToDefaultView
            Write-Host "[OK] Added Boolean field: $($field.Name)" -ForegroundColor Green
        }
    }
}

# Create folders for organization
$folders = @("Word", "Excel", "PowerPoint", "Uploads")
foreach ($folder in $folders) {
    $existingFolder = Get-PnPFolder -Url "$LibraryName/$folder" -ErrorAction SilentlyContinue
    if ($existingFolder) {
        Write-Host "[SKIP] Folder '$folder' already exists." -ForegroundColor Yellow
    } else {
        Add-PnPFolder -Name $folder -Folder $LibraryName
        Write-Host "[OK] Created folder: $folder" -ForegroundColor Green
    }
}

Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " $LibraryName provisioning complete!" -ForegroundColor Green
Write-Host "============================================`n" -ForegroundColor Cyan
Write-Host "Library URL: $SiteUrl/$LibraryName" -ForegroundColor White
