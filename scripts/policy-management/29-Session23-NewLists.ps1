# ============================================================================
# 29-Session23-NewLists.ps1
# Creates all new SharePoint lists added in Session 23
# Includes: SecurityAuditLog, SecurityAlerts, SyncLog, SyncConfig,
#           LegalHolds, Audiences, Delegations, ReminderSchedule, EventLog
# ============================================================================
# PREREQUISITE: Already connected to SharePoint via Connect-PnPOnline
# Site: https://mf7m.sharepoint.com/sites/PolicyManager
# ============================================================================

$siteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"

# ─── Helper: Create list if it doesn't exist ───
function Ensure-List {
    param([string]$Title, [string]$Description)
    $existing = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  ✓ $Title already exists ($($existing.ItemCount) items)" -ForegroundColor Green
        return $false
    }
    New-PnPList -Title $Title -Template GenericList -EnableVersioning -OnQuickLaunch:$false
    Set-PnPList -Identity $Title -Description $Description
    Write-Host "  ✓ $Title created" -ForegroundColor Cyan
    return $true
}

# ─── Helper: Add field if it doesn't exist ───
function Ensure-Field {
    param([string]$List, [string]$Name, [string]$Type, [string]$Choices = "", [bool]$Required = $false)
    $existing = Get-PnPField -List $List -Identity $Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "    - $Name already exists" -ForegroundColor DarkGray
        return
    }
    switch ($Type) {
        "Text"      { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Text -AddToDefaultView }
        "Note"      { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Note -AddToDefaultView }
        "Number"    { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Number -AddToDefaultView }
        "DateTime"  { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type DateTime -AddToDefaultView }
        "Boolean"   { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Boolean -AddToDefaultView }
        "Choice"    {
            $choiceArr = $Choices -split ";"
            Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Choice -Choices $choiceArr -AddToDefaultView
        }
    }
    Write-Host "    + $Name ($Type)" -ForegroundColor Green
}

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Session 23 — New SharePoint Lists & Columns" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Cyan

# ═══════════════════════════════════════════════════
# 1. PM_SecurityAuditLog — Security events with risk scoring
# ═══════════════════════════════════════════════════
Write-Host "[1/9] PM_SecurityAuditLog" -ForegroundColor Yellow
Ensure-List -Title "PM_SecurityAuditLog" -Description "Security audit events with risk scoring and threat detection"
Ensure-Field -List "PM_SecurityAuditLog" -Name "EventType" -Type "Text"
Ensure-Field -List "PM_SecurityAuditLog" -Name "Severity" -Type "Choice" -Choices "Low;Medium;High;Critical"
Ensure-Field -List "PM_SecurityAuditLog" -Name "UserEmail" -Type "Text"
Ensure-Field -List "PM_SecurityAuditLog" -Name "UserDisplayName" -Type "Text"
Ensure-Field -List "PM_SecurityAuditLog" -Name "AttemptedRole" -Type "Text"
Ensure-Field -List "PM_SecurityAuditLog" -Name "ActualRoles" -Type "Text"
Ensure-Field -List "PM_SecurityAuditLog" -Name "AttemptedApp" -Type "Text"
Ensure-Field -List "PM_SecurityAuditLog" -Name "IPAddress" -Type "Text"
Ensure-Field -List "PM_SecurityAuditLog" -Name "UserAgent" -Type "Note"
Ensure-Field -List "PM_SecurityAuditLog" -Name "Details" -Type "Note"
Ensure-Field -List "PM_SecurityAuditLog" -Name "RiskScore" -Type "Number"
Ensure-Field -List "PM_SecurityAuditLog" -Name "SessionId" -Type "Text"
Ensure-Field -List "PM_SecurityAuditLog" -Name "AuditTimestamp" -Type "DateTime"

# ═══════════════════════════════════════════════════
# 2. PM_SecurityAlerts — Active security alerts
# ═══════════════════════════════════════════════════
Write-Host "`n[2/9] PM_SecurityAlerts" -ForegroundColor Yellow
Ensure-List -Title "PM_SecurityAlerts" -Description "Active security alerts and threat detection"
Ensure-Field -List "PM_SecurityAlerts" -Name "Description" -Type "Note"
Ensure-Field -List "PM_SecurityAlerts" -Name "Severity" -Type "Choice" -Choices "Low;Medium;High;Critical"
Ensure-Field -List "PM_SecurityAlerts" -Name "Category" -Type "Choice" -Choices "access;data;identity;compliance;threat;policy"
Ensure-Field -List "PM_SecurityAlerts" -Name "RiskScore" -Type "Number"
Ensure-Field -List "PM_SecurityAlerts" -Name "AffectedUsers" -Type "Number"
Ensure-Field -List "PM_SecurityAlerts" -Name "AlertStatus" -Type "Choice" -Choices "new;investigating;resolved;dismissed"
Ensure-Field -List "PM_SecurityAlerts" -Name "AlertTimestamp" -Type "DateTime"

# ═══════════════════════════════════════════════════
# 3. PM_SyncLog — EntraID sync operation history
# ═══════════════════════════════════════════════════
Write-Host "`n[3/9] PM_SyncLog" -ForegroundColor Yellow
Ensure-List -Title "PM_SyncLog" -Description "EntraID sync operation history"
Ensure-Field -List "PM_SyncLog" -Name "SyncId" -Type "Text"
Ensure-Field -List "PM_SyncLog" -Name "SyncType" -Type "Choice" -Choices "Full;Delta;Filtered;Single"
Ensure-Field -List "PM_SyncLog" -Name "Status" -Type "Choice" -Choices "Running;Completed;CompletedWithErrors;Failed"
Ensure-Field -List "PM_SyncLog" -Name "Message" -Type "Note"
Ensure-Field -List "PM_SyncLog" -Name "UsersProcessed" -Type "Number"
Ensure-Field -List "PM_SyncLog" -Name "UsersAdded" -Type "Number"
Ensure-Field -List "PM_SyncLog" -Name "UsersUpdated" -Type "Number"
Ensure-Field -List "PM_SyncLog" -Name "UsersDeactivated" -Type "Number"
Ensure-Field -List "PM_SyncLog" -Name "ErrorCount" -Type "Number"
Ensure-Field -List "PM_SyncLog" -Name "Duration" -Type "Number"
Ensure-Field -List "PM_SyncLog" -Name "TriggeredBy" -Type "Text"

# ═══════════════════════════════════════════════════
# 4. PM_SyncConfig — Sync configuration and delta tokens
# ═══════════════════════════════════════════════════
Write-Host "`n[4/9] PM_SyncConfig" -ForegroundColor Yellow
Ensure-List -Title "PM_SyncConfig" -Description "Sync configuration and delta tokens"
Ensure-Field -List "PM_SyncConfig" -Name "ConfigType" -Type "Text"
Ensure-Field -List "PM_SyncConfig" -Name "ConfigValue" -Type "Note"

# ═══════════════════════════════════════════════════
# 5. PM_LegalHolds — Legal hold records
# ═══════════════════════════════════════════════════
Write-Host "`n[5/9] PM_LegalHolds" -ForegroundColor Yellow
Ensure-List -Title "PM_LegalHolds" -Description "Legal hold records for compliance locks"
Ensure-Field -List "PM_LegalHolds" -Name "PolicyId" -Type "Number"
Ensure-Field -List "PM_LegalHolds" -Name "PolicyName" -Type "Text"
Ensure-Field -List "PM_LegalHolds" -Name "HoldReason" -Type "Note"
Ensure-Field -List "PM_LegalHolds" -Name "CaseReference" -Type "Text"
Ensure-Field -List "PM_LegalHolds" -Name "PlacedByEmail" -Type "Text"
Ensure-Field -List "PM_LegalHolds" -Name "PlacedByName" -Type "Text"
Ensure-Field -List "PM_LegalHolds" -Name "PlacedDate" -Type "DateTime"
Ensure-Field -List "PM_LegalHolds" -Name "ExpiryDate" -Type "DateTime"
Ensure-Field -List "PM_LegalHolds" -Name "ReleasedByEmail" -Type "Text"
Ensure-Field -List "PM_LegalHolds" -Name "ReleasedDate" -Type "DateTime"
Ensure-Field -List "PM_LegalHolds" -Name "ReleaseReason" -Type "Note"
Ensure-Field -List "PM_LegalHolds" -Name "Status" -Type "Choice" -Choices "Active;Released;Expired"

# ═══════════════════════════════════════════════════
# 6. PM_Audiences — Audience targeting rules
# ═══════════════════════════════════════════════════
Write-Host "`n[6/9] PM_Audiences" -ForegroundColor Yellow
Ensure-List -Title "PM_Audiences" -Description "Audience targeting rules and member counts"
Ensure-Field -List "PM_Audiences" -Name "AudienceName" -Type "Text"
Ensure-Field -List "PM_Audiences" -Name "Description" -Type "Note"
Ensure-Field -List "PM_Audiences" -Name "Rules" -Type "Note"
Ensure-Field -List "PM_Audiences" -Name "RuleOperator" -Type "Choice" -Choices "AND;OR"
Ensure-Field -List "PM_Audiences" -Name "MemberCount" -Type "Number"
Ensure-Field -List "PM_Audiences" -Name "LastEvaluated" -Type "DateTime"
Ensure-Field -List "PM_Audiences" -Name "IsActive" -Type "Boolean"

# ═══════════════════════════════════════════════════
# 7. PM_Delegations — Approval and review delegations
# ═══════════════════════════════════════════════════
Write-Host "`n[7/9] PM_Delegations" -ForegroundColor Yellow
Ensure-List -Title "PM_Delegations" -Description "Approval and review delegations"
Ensure-Field -List "PM_Delegations" -Name "DelegatedById" -Type "Number"
Ensure-Field -List "PM_Delegations" -Name "DelegatedByEmail" -Type "Text"
Ensure-Field -List "PM_Delegations" -Name "DelegatedToId" -Type "Number"
Ensure-Field -List "PM_Delegations" -Name "DelegatedToEmail" -Type "Text"
Ensure-Field -List "PM_Delegations" -Name "DelegationType" -Type "Choice" -Choices "Approval;Review;Both"
Ensure-Field -List "PM_Delegations" -Name "Reason" -Type "Text"
Ensure-Field -List "PM_Delegations" -Name "StartDate" -Type "DateTime"
Ensure-Field -List "PM_Delegations" -Name "EndDate" -Type "DateTime"
Ensure-Field -List "PM_Delegations" -Name "Status" -Type "Choice" -Choices "Active;Expired;Revoked"

# ═══════════════════════════════════════════════════
# 8. PM_ReminderSchedule — Scheduled reminders
# ═══════════════════════════════════════════════════
Write-Host "`n[8/9] PM_ReminderSchedule" -ForegroundColor Yellow
Ensure-List -Title "PM_ReminderSchedule" -Description "Scheduled acknowledgement reminders"
Ensure-Field -List "PM_ReminderSchedule" -Name "PolicyId" -Type "Number"
Ensure-Field -List "PM_ReminderSchedule" -Name "PolicyName" -Type "Text"
Ensure-Field -List "PM_ReminderSchedule" -Name "RecipientId" -Type "Number"
Ensure-Field -List "PM_ReminderSchedule" -Name "RecipientEmail" -Type "Text"
Ensure-Field -List "PM_ReminderSchedule" -Name "ReminderType" -Type "Choice" -Choices "3-Day;1-Day;Overdue;Custom"
Ensure-Field -List "PM_ReminderSchedule" -Name "ScheduledDate" -Type "DateTime"
Ensure-Field -List "PM_ReminderSchedule" -Name "SentDate" -Type "DateTime"
Ensure-Field -List "PM_ReminderSchedule" -Name "Status" -Type "Choice" -Choices "Scheduled;Sent;Skipped;Failed"

# ═══════════════════════════════════════════════════
# 9. PM_EventLog — Event Viewer diagnostic events
# ═══════════════════════════════════════════════════
Write-Host "`n[9/9] PM_EventLog" -ForegroundColor Yellow
Ensure-List -Title "PM_EventLog" -Description "Event Viewer diagnostic events"
Ensure-Field -List "PM_EventLog" -Name "EventCode" -Type "Text"
Ensure-Field -List "PM_EventLog" -Name "EventCategory" -Type "Text"
Ensure-Field -List "PM_EventLog" -Name "Severity" -Type "Choice" -Choices "Debug;Info;Warning;Error;Critical"
Ensure-Field -List "PM_EventLog" -Name "Source" -Type "Text"
Ensure-Field -List "PM_EventLog" -Name "Message" -Type "Note"
Ensure-Field -List "PM_EventLog" -Name "Details" -Type "Note"
Ensure-Field -List "PM_EventLog" -Name "UserEmail" -Type "Text"
Ensure-Field -List "PM_EventLog" -Name "PageUrl" -Type "Text"
Ensure-Field -List "PM_EventLog" -Name "SessionId" -Type "Text"
Ensure-Field -List "PM_EventLog" -Name "EventTimestamp" -Type "DateTime"

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  All 9 lists processed!" -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Green
