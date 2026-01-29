# ============================================================================
# DWx Policy Manager - Notification Lists
# Part 7: PM_NotificationQueue, PM_Notifications
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,

    [Parameter(Mandatory=$false)]
    [switch]$UseWebLogin
)

# Connect to SharePoint
Write-Host "Connecting to SharePoint: $SiteUrl" -ForegroundColor Cyan
if ($UseWebLogin) {
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
} else {
    Connect-PnPOnline -Url $SiteUrl -Interactive
}

# ============================================================================
# LIST 21: PM_NotificationQueue
# Queue for outbound policy notifications (email, Teams, in-app)
# ============================================================================
Write-Host "`n Creating PM_NotificationQueue list..." -ForegroundColor Yellow

$listName = "PM_NotificationQueue"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Notification Type" -InternalName "NotificationType" -Type Choice -Choices "PolicyShared","PolicyFollowed","PolicyUpdated","PolicyAcknowledgmentRequired","PolicyAcknowledged","PolicyExpiring","PolicyPublished","PolicyComment","Custom" -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Recipient Email" -InternalName "RecipientEmail" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Recipient User ID" -InternalName "RecipientUserId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Recipient Name" -InternalName "RecipientName" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Sender Email" -InternalName "SenderEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Sender User ID" -InternalName "SenderUserId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Sender Name" -InternalName "SenderName" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Title" -InternalName "PolicyTitle" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Version" -InternalName "PolicyVersion" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Message" -InternalName "Message" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Channel" -InternalName "Channel" -Type Choice -Choices "Email","Teams","InApp","All" -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices "Low","Normal","High","Urgent" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Pending","Processing","Sent","Failed","Retry" -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Retry Count" -InternalName "RetryCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Max Retries" -InternalName "MaxRetries" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Last Error" -InternalName "LastError" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Scheduled Send Time" -InternalName "ScheduledSendTime" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Sent Time" -InternalName "SentTime" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Related Share ID" -InternalName "RelatedShareId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Related Follow ID" -InternalName "RelatedFollowId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Teams Channel ID" -InternalName "TeamsChannelId" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Teams Team ID" -InternalName "TeamsTeamId" -Type Text -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "Status" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "NotificationType" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "RecipientEmail" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "Priority" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_NotificationQueue list configured" -ForegroundColor Green

# ============================================================================
# LIST 22: PM_Notifications
# In-app notifications displayed to users within the Policy Manager UI
# ============================================================================
Write-Host "`n Creating PM_Notifications list..." -ForegroundColor Yellow

$listName = "PM_Notifications"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Message" -InternalName "Message" -Type Note -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Recipient" -InternalName "RecipientId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Type" -InternalName "Type" -Type Choice -Choices "PolicyShare","PolicyFollow","PolicyUpdate","PolicyAcknowledgment","PolicyExpiring","Policy" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices "Low","Normal","High","Urgent" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Read" -InternalName "IsRead" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Related Item Type" -InternalName "RelatedItemType" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Related Item ID" -InternalName "RelatedItemId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Action URL" -InternalName "ActionUrl" -Type Text -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "RecipientId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsRead" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "Type" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_Notifications list configured" -ForegroundColor Green

Write-Host "`n Notification lists created successfully!" -ForegroundColor Green
Write-Host "   - PM_NotificationQueue" -ForegroundColor White
Write-Host "   - PM_Notifications" -ForegroundColor White
