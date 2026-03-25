# ============================================================================
# Script 20: Update Notification Choice Fields
# Removes and recreates Type/NotificationType fields with expanded choices
# Idempotent — safe to run multiple times
# ============================================================================

Write-Host "`n=== Updating Notification Choice Fields ===" -ForegroundColor Cyan

# PM_Notifications — Type field
Write-Host "`n Updating PM_Notifications.Type choices..." -ForegroundColor Yellow
try {
    Remove-PnPField -List "PM_Notifications" -Identity "Type" -Force -ErrorAction SilentlyContinue
    Add-PnPField -List "PM_Notifications" -DisplayName "Type" -InternalName "Type" -Type Choice -Choices "Policy","PolicyShare","PolicyFollow","PolicyUpdate","PolicyAcknowledgment","PolicyExpiring","ApprovalRequired","ApprovalComplete","ApprovalRejected","ApprovalDelegated","ApprovalCancelled","ApprovalExpired","ApprovalEscalated","ReviewRequired","ReviewCancelled","Reminder","Nudge" -AddToDefaultView -ErrorAction Stop
    Set-PnPField -List "PM_Notifications" -Identity "Type" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
    Write-Host "  PM_Notifications.Type updated successfully" -ForegroundColor Green
} catch {
    Write-Host "  Failed: $_" -ForegroundColor Red
}

# PM_NotificationQueue — NotificationType field
Write-Host "`n Updating PM_NotificationQueue.NotificationType choices..." -ForegroundColor Yellow
try {
    Remove-PnPField -List "PM_NotificationQueue" -Identity "NotificationType" -Force -ErrorAction SilentlyContinue
    Add-PnPField -List "PM_NotificationQueue" -DisplayName "Notification Type" -InternalName "NotificationType" -Type Choice -Choices "PolicyShared","PolicyFollowed","PolicyUpdated","PolicyAcknowledgmentRequired","PolicyAcknowledged","PolicyExpiring","PolicyPublished","PolicyComment","Custom","ReviewRequired","ReviewCancelled","ApprovalRequired","ApprovalComplete","ApprovalRejected","review-due","approval-request","policy-published","ack-required","sla-breach" -AddToDefaultView -ErrorAction Stop
    Write-Host "  PM_NotificationQueue.NotificationType updated successfully" -ForegroundColor Green
} catch {
    Write-Host "  Failed: $_" -ForegroundColor Red
}

Write-Host "`n=== Notification Choice Update Complete ===" -ForegroundColor Cyan
