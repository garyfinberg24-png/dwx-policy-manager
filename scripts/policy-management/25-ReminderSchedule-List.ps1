# ============================================================================
# Policy Manager — PM_ReminderSchedule List
# Tracks scheduled reminders for policy revisions, acknowledgement deadlines,
# and recurring review cycles.
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  PM_ReminderSchedule List Provisioning" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$listName = "PM_ReminderSchedule"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

# Core fields
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Policy Title" -InternalName "PolicyTitle" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Reminder Type" -InternalName "ReminderType" -Type Choice -Choices "RevisionDue","AcknowledgementOverdue","ReviewCycleDue","ExpiryWarning","CustomReminder" -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Scheduled Date" -InternalName "ScheduledDate" -Type DateTime -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Reminder Status" -InternalName "ReminderStatus" -Type Choice -Choices "Pending","Sent","Skipped","Failed" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "ReminderStatus" -Values @{DefaultValue="Pending"} -ErrorAction SilentlyContinue

# Recipient targeting
Add-PnPField -List $listName -DisplayName "Recipient Type" -InternalName "RecipientType" -Type Choice -Choices "Author","Reviewer","Approver","AllAssigned","Custom" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Recipient Email" -InternalName "RecipientEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Recipient Id" -InternalName "RecipientId" -Type Number -ErrorAction SilentlyContinue | Out-Null

# Recurrence
Add-PnPField -List $listName -DisplayName "Is Recurring" -InternalName "IsRecurring" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Recurrence Interval" -InternalName "RecurrenceInterval" -Type Choice -Choices "Daily","Weekly","Monthly","Quarterly","Annual" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Next Occurrence" -InternalName "NextOccurrence" -Type DateTime -ErrorAction SilentlyContinue | Out-Null

# Tracking
Add-PnPField -List $listName -DisplayName "Sent Date" -InternalName "SentDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Failure Reason" -InternalName "FailureReason" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Created By Email" -InternalName "CreatedByEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null

# Indexes
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "ReminderStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "ScheduledDate" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  PM_ReminderSchedule list provisioned successfully" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
