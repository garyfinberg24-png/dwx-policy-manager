# ============================================================================
# Policy Manager — Retention & Legal Holds Lists
# Provisions PM_RetentionPolicies, PM_LegalHolds, PM_RetentionArchive
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Retention & Legal Holds List Provisioning" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ─── PM_RetentionPolicies ────────────────────────────────────────────

$listName = "PM_RetentionPolicies"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Entity Type" -InternalName "EntityType" -Type Choice -Choices "Policy","Document","AuditLog","Quiz" -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Retention Days" -InternalName "RetentionDays" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Retention Action" -InternalName "RetentionAction" -Type Choice -Choices "Archive","Delete","Review" -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Created By" -InternalName "CreatedByUser" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Created Date" -InternalName "CreatedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null

# Indexes
Set-PnPField -List $listName -Identity "EntityType" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  Fields added to $listName" -ForegroundColor Green
Write-Host ""

# ─── PM_LegalHolds ───────────────────────────────────────────────────

$listName = "PM_LegalHolds"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy Id" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Policy Title" -InternalName "PolicyTitle" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Hold Reason" -InternalName "HoldReason" -Type Note -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Placed By" -InternalName "PlacedBy" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Placed By Email" -InternalName "PlacedByEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Placed Date" -InternalName "PlacedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Expiry Date" -InternalName "ExpiryDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Active","Released","Expired" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "Status" -Values @{DefaultValue="Active"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Released By" -InternalName "ReleasedBy" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Released Date" -InternalName "ReleasedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Release Reason" -InternalName "ReleaseReason" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Case Reference" -InternalName "CaseReference" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Compliance Relevant" -InternalName "ComplianceRelevant" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "ComplianceRelevant" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue

# Indexes
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "Status" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  Fields added to $listName" -ForegroundColor Green
Write-Host ""

# ─── PM_RetentionArchive ─────────────────────────────────────────────

$listName = "PM_RetentionArchive"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Original Policy Id" -InternalName "OriginalPolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Policy Title" -InternalName "PolicyTitle" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Policy Number" -InternalName "PolicyNumber" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Policy Category" -InternalName "PolicyCategory" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Archived Date" -InternalName "ArchivedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Archived By" -InternalName "ArchivedBy" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Retention Rule Id" -InternalName "RetentionRuleId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Original Content" -InternalName "OriginalContent" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Original Metadata" -InternalName "OriginalMetadata" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Archive Reason" -InternalName "ArchiveReason" -Type Text -ErrorAction SilentlyContinue | Out-Null

# Indexes
Set-PnPField -List $listName -Identity "OriginalPolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "ArchivedDate" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  Fields added to $listName" -ForegroundColor Green
Write-Host ""

Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Retention & Legal Holds lists provisioned successfully" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
