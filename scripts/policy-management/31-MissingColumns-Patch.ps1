# ============================================================================
# 31-MissingColumns-Patch.ps1
# Adds missing columns to existing lists that were identified during
# Session 23 audit. Idempotent — safe to run multiple times.
# ============================================================================
# PREREQUISITE: Already connected to SharePoint via Connect-PnPOnline
# ============================================================================

function Ensure-ColumnOnList {
    param([string]$List, [string]$Name, [string]$Type, [string]$Choices = "")
    $existing = Get-PnPField -List $List -Identity $Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "    - $List.$Name already exists" -ForegroundColor DarkGray
        return
    }
    switch ($Type) {
        "Text"      { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Text }
        "Note"      { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Note }
        "Number"    { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Number }
        "DateTime"  { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type DateTime }
        "Boolean"   { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Boolean }
        "Choice"    {
            $choiceArr = $Choices -split ";"
            Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Choice -Choices $choiceArr
        }
    }
    Write-Host "    + $List.$Name ($Type)" -ForegroundColor Green
}

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Missing Columns Patch — Session 23" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Cyan

# ─── PM_UserProfiles — PMRoles column for multi-role support ───
Write-Host "[1] PM_UserProfiles — PMRoles" -ForegroundColor Yellow
Ensure-ColumnOnList -List "PM_UserProfiles" -Name "PMRoles" -Type "Text"
Ensure-ColumnOnList -List "PM_UserProfiles" -Name "EntraObjectId" -Type "Text"
Ensure-ColumnOnList -List "PM_UserProfiles" -Name "LastSyncedAt" -Type "DateTime"

# ─── PM_PolicyTemplates — Compliance & Metadata fields ───
Write-Host "`n[2] PM_PolicyTemplates — Compliance fields" -ForegroundColor Yellow
Ensure-ColumnOnList -List "PM_PolicyTemplates" -Name "ComplianceRisk" -Type "Choice" -Choices "Critical;High;Medium;Low;Informational"
Ensure-ColumnOnList -List "PM_PolicyTemplates" -Name "SuggestedReadTimeframe" -Type "Choice" -Choices "Immediate;Day1;Day3;Week1;Week2;Month1;Month3;Month6;Custom"
Ensure-ColumnOnList -List "PM_PolicyTemplates" -Name "RequiresAcknowledgement" -Type "Boolean"
Ensure-ColumnOnList -List "PM_PolicyTemplates" -Name "RequiresQuiz" -Type "Boolean"
Ensure-ColumnOnList -List "PM_PolicyTemplates" -Name "KeyPointsTemplate" -Type "Note"
Ensure-ColumnOnList -List "PM_PolicyTemplates" -Name "RegulatoryFramework" -Type "Text"
Ensure-ColumnOnList -List "PM_PolicyTemplates" -Name "RegulatoryReferences" -Type "Note"

# ─── PM_Policies — MetadataProfileId (optional) ───
Write-Host "`n[3] PM_Policies — Additional columns" -ForegroundColor Yellow
Ensure-ColumnOnList -List "PM_Policies" -Name "MetadataProfileId" -Type "Number"
Ensure-ColumnOnList -List "PM_Policies" -Name "HTMLContent" -Type "Note"

# ─── PM_PolicyCategories — IsDefault column ───
Write-Host "`n[4] PM_PolicyCategories — IsDefault" -ForegroundColor Yellow
Ensure-ColumnOnList -List "PM_PolicyCategories" -Name "IsDefault" -Type "Boolean"

# ─── PM_Configuration — Navigation visibility (already via service, but ensure column exists) ───
Write-Host "`n[5] PM_Configuration — Verify core columns" -ForegroundColor Yellow
Ensure-ColumnOnList -List "PM_Configuration" -Name "ConfigKey" -Type "Text"
Ensure-ColumnOnList -List "PM_Configuration" -Name "ConfigValue" -Type "Note"
Ensure-ColumnOnList -List "PM_Configuration" -Name "Category" -Type "Text"
Ensure-ColumnOnList -List "PM_Configuration" -Name "IsActive" -Type "Boolean"
Ensure-ColumnOnList -List "PM_Configuration" -Name "IsSystemConfig" -Type "Boolean"

# ─── PM_NotificationQueue — Ensure all required columns ───
Write-Host "`n[6] PM_NotificationQueue — Notification columns" -ForegroundColor Yellow
Ensure-ColumnOnList -List "PM_NotificationQueue" -Name "To" -Type "Text"
Ensure-ColumnOnList -List "PM_NotificationQueue" -Name "RecipientEmail" -Type "Text"
Ensure-ColumnOnList -List "PM_NotificationQueue" -Name "Subject" -Type "Text"
Ensure-ColumnOnList -List "PM_NotificationQueue" -Name "Message" -Type "Note"
Ensure-ColumnOnList -List "PM_NotificationQueue" -Name "QueueStatus" -Type "Choice" -Choices "Pending;Processing;Sent;Failed"
Ensure-ColumnOnList -List "PM_NotificationQueue" -Name "Priority" -Type "Choice" -Choices "Low;Normal;High;Urgent"
Ensure-ColumnOnList -List "PM_NotificationQueue" -Name "NotificationType" -Type "Text"
Ensure-ColumnOnList -List "PM_NotificationQueue" -Name "Channel" -Type "Choice" -Choices "Email;Teams;InApp"
Ensure-ColumnOnList -List "PM_NotificationQueue" -Name "RetryCount" -Type "Number"
Ensure-ColumnOnList -List "PM_NotificationQueue" -Name "ErrorMessage" -Type "Note"

# ─── PM_PolicyAuditLog — Ensure all columns for enhanced security audit ───
Write-Host "`n[7] PM_PolicyAuditLog — Security audit columns" -ForegroundColor Yellow
Ensure-ColumnOnList -List "PM_PolicyAuditLog" -Name "AuditAction" -Type "Text"
Ensure-ColumnOnList -List "PM_PolicyAuditLog" -Name "EntityType" -Type "Text"
Ensure-ColumnOnList -List "PM_PolicyAuditLog" -Name "EntityId" -Type "Text"
Ensure-ColumnOnList -List "PM_PolicyAuditLog" -Name "PolicyId" -Type "Number"
Ensure-ColumnOnList -List "PM_PolicyAuditLog" -Name "ActionDescription" -Type "Note"
Ensure-ColumnOnList -List "PM_PolicyAuditLog" -Name "PerformedByEmail" -Type "Text"
Ensure-ColumnOnList -List "PM_PolicyAuditLog" -Name "PerformedById" -Type "Number"
Ensure-ColumnOnList -List "PM_PolicyAuditLog" -Name "ComplianceRelevant" -Type "Boolean"
Ensure-ColumnOnList -List "PM_PolicyAuditLog" -Name "ActionDate" -Type "DateTime"

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  All missing columns patched!" -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Green
