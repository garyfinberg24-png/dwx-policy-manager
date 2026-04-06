# Patch-NotificationQueueColumns.ps1
# Adds missing columns to PM_NotificationQueue that the Logic App and code expect.
# Idempotent — safe to run multiple times.
# Assumes you are already connected via Connect-PnPOnline.

$listName = "PM_NotificationQueue"

Write-Host "Patching $listName with missing columns..." -ForegroundColor Cyan

# 'To' — Logic App reads this as the email recipient
$field = Get-PnPField -List $listName -Identity "To" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $listName -DisplayName "To" -InternalName "To" -Type Text -ErrorAction SilentlyContinue
    Write-Host "  + Added 'To' column" -ForegroundColor Green
} else {
    Write-Host "  - 'To' already exists" -ForegroundColor Gray
}

# 'Subject' — Logic App reads this as the email subject line
$field = Get-PnPField -List $listName -Identity "Subject" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $listName -DisplayName "Subject" -InternalName "Subject" -Type Text -ErrorAction SilentlyContinue
    Write-Host "  + Added 'Subject' column" -ForegroundColor Green
} else {
    Write-Host "  - 'Subject' already exists" -ForegroundColor Gray
}

# 'QueueStatus' — code writes status here; Logic App filters on this
$field = Get-PnPField -List $listName -Identity "QueueStatus" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $listName -DisplayName "Queue Status" -InternalName "QueueStatus" -Type Choice -Choices "Pending","Processing","Sent","Failed","Retry" -ErrorAction SilentlyContinue
    Set-PnPField -List $listName -Identity "QueueStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
    Write-Host "  + Added 'QueueStatus' column (indexed)" -ForegroundColor Green
} else {
    Write-Host "  - 'QueueStatus' already exists" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Patch complete. Verify with: .\scripts\Verify-PublishPipeline.ps1" -ForegroundColor Green
