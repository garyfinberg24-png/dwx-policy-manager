# ============================================================================
# 34-Escalation-Columns.ps1
# Adds escalation tracking columns to PM_Approvals
# ============================================================================
# PREREQUISITE: Already connected to SharePoint via Connect-PnPOnline
# ============================================================================

$listName = "PM_Approvals"

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Adding Escalation Columns to $listName" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Cyan

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "  ✗ $listName does not exist" -ForegroundColor Red
    return
}

# EscalationCount — Number
$field = Get-PnPField -List $listName -Identity "EscalationCount" -ErrorAction SilentlyContinue
if (-not $field) {
    Add-PnPField -List $listName -DisplayName "EscalationCount" -InternalName "EscalationCount" -Type Number
    Set-PnPField -List $listName -Identity "EscalationCount" -Values @{DefaultValue="0"}
    Write-Host "  + EscalationCount (Number, default 0)" -ForegroundColor Green
} else {
    Write-Host "  - EscalationCount already exists" -ForegroundColor DarkGray
}

# Comments — Note (may already exist)
$field = Get-PnPField -List $listName -Identity "Comments" -ErrorAction SilentlyContinue
if (-not $field) {
    Add-PnPField -List $listName -DisplayName "Comments" -InternalName "Comments" -Type Note
    Write-Host "  + Comments (Note)" -ForegroundColor Green
} else {
    Write-Host "  - Comments already exists" -ForegroundColor DarkGray
}

# ApprovalLevel — Text (may already exist)
$field = Get-PnPField -List $listName -Identity "ApprovalLevel" -ErrorAction SilentlyContinue
if (-not $field) {
    Add-PnPField -List $listName -DisplayName "ApprovalLevel" -InternalName "ApprovalLevel" -Type Text
    Write-Host "  + ApprovalLevel (Text)" -ForegroundColor Green
} else {
    Write-Host "  - ApprovalLevel already exists" -ForegroundColor DarkGray
}

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  Escalation columns complete!" -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Green
