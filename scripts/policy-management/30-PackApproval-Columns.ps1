# ============================================================================
# 30-PackApproval-Columns.ps1
# Adds approval workflow columns to PM_PolicyPacks
# ============================================================================
# PREREQUISITE: Already connected to SharePoint via Connect-PnPOnline
# ============================================================================

$listName = "PM_PolicyPacks"

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Adding Approval Columns to $listName" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Cyan

# Check list exists
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "  ✗ $listName does not exist — run provisioning first" -ForegroundColor Red
    return
}

# ApprovalStatus — Choice field
$field = Get-PnPField -List $listName -Identity "ApprovalStatus" -ErrorAction SilentlyContinue
if (-not $field) {
    Add-PnPField -List $listName -DisplayName "ApprovalStatus" -InternalName "ApprovalStatus" -Type Choice -Choices "Draft","Pending Approval","Approved","Rejected","Changes Requested" -AddToDefaultView
    Set-PnPField -List $listName -Identity "ApprovalStatus" -Values @{DefaultValue="Draft"}
    Write-Host "  + ApprovalStatus (Choice: Draft|Pending Approval|Approved|Rejected|Changes Requested)" -ForegroundColor Green
} else {
    Write-Host "  - ApprovalStatus already exists" -ForegroundColor DarkGray
}

# ApproverEmails — Text field (semicolon-separated)
$field = Get-PnPField -List $listName -Identity "ApproverEmails" -ErrorAction SilentlyContinue
if (-not $field) {
    Add-PnPField -List $listName -DisplayName "ApproverEmails" -InternalName "ApproverEmails" -Type Note -AddToDefaultView
    Write-Host "  + ApproverEmails (Note — semicolon-separated email list)" -ForegroundColor Green
} else {
    Write-Host "  - ApproverEmails already exists" -ForegroundColor DarkGray
}

# ApprovedByEmail — Text field
$field = Get-PnPField -List $listName -Identity "ApprovedByEmail" -ErrorAction SilentlyContinue
if (-not $field) {
    Add-PnPField -List $listName -DisplayName "ApprovedByEmail" -InternalName "ApprovedByEmail" -Type Text -AddToDefaultView
    Write-Host "  + ApprovedByEmail (Text)" -ForegroundColor Green
} else {
    Write-Host "  - ApprovedByEmail already exists" -ForegroundColor DarkGray
}

# ApprovedDate — DateTime field
$field = Get-PnPField -List $listName -Identity "ApprovedDate" -ErrorAction SilentlyContinue
if (-not $field) {
    Add-PnPField -List $listName -DisplayName "ApprovedDate" -InternalName "ApprovedDate" -Type DateTime -AddToDefaultView
    Write-Host "  + ApprovedDate (DateTime)" -ForegroundColor Green
} else {
    Write-Host "  - ApprovedDate already exists" -ForegroundColor DarkGray
}

# ApprovalComments — Note field
$field = Get-PnPField -List $listName -Identity "ApprovalComments" -ErrorAction SilentlyContinue
if (-not $field) {
    Add-PnPField -List $listName -DisplayName "ApprovalComments" -InternalName "ApprovalComments" -Type Note -AddToDefaultView
    Write-Host "  + ApprovalComments (Note)" -ForegroundColor Green
} else {
    Write-Host "  - ApprovalComments already exists" -ForegroundColor DarkGray
}

# CreatedByEmail — Text field (for notification back to pack creator)
$field = Get-PnPField -List $listName -Identity "CreatedByEmail" -ErrorAction SilentlyContinue
if (-not $field) {
    Add-PnPField -List $listName -DisplayName "CreatedByEmail" -InternalName "CreatedByEmail" -Type Text
    Write-Host "  + CreatedByEmail (Text)" -ForegroundColor Green
} else {
    Write-Host "  - CreatedByEmail already exists" -ForegroundColor DarkGray
}

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  PM_PolicyPacks approval columns complete!" -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Green
